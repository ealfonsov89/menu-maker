#[macro_use] extern crate log;
extern crate simplelog;
use simplelog::*;

use std::fs::File;


use calamine::Data;
use calamine::Range;
use serde::Serialize;
use std::env;
use std::fs;
use std::path::Path;
use std::process::Command;
use std::env::current_dir;
use tera::Context;
use tera::Tera;
use calamine::{open_workbook, Xlsx, Reader, RangeDeserializerBuilder};
use chrono::Local;

#[derive(Serialize)]
struct Item {
    name: String,
    price: String,
}

fn main() {
    prepare_log();


    let exe_dir = env::current_exe()
        .unwrap()
        .parent()
        .unwrap()
        .to_path_buf();

    // Cambiar directorio de trabajo
    env::set_current_dir(&exe_dir).unwrap();
    info!("Directorio actual: {}", current_dir().unwrap().display());
    let template_path = format!("{}/template/*.html", exe_dir.display()).trim().replace('\\', "/");
    info!("Usando plantillas en: {}", template_path);

    

    let file_path = match handle_arguments() {
        Some(value) => value,
        None => return,
    };
    match Tera::new(template_path.as_str()) {
        Ok(tera) => process(tera, file_path.to_string()),
        Err(e) => {
            error!("Error parsing templates: {}", e);
            ::std::process::exit(1);
        }
    };
}

fn prepare_log() {
    CombinedLogger::init(vec![
        TermLogger::new(LevelFilter::Info, Config::default(), TerminalMode::Mixed, ColorChoice::Auto),
        WriteLogger::new(LevelFilter::Info, Config::default(), File::create("output.log").unwrap()),
    ]).unwrap();
}

fn handle_arguments() -> Option<String> {
    let args: Vec<String> = env::args().collect();
    let mut file_path = String::new();
    if args.len() < 2 {
        println!("Agrega la ruta al xlsx:");
        match std::io::stdin().read_line(&mut file_path) {
            Ok(_) => {
                info!("Archivo recibido desde stdin: {}", file_path.trim());
            },
            Err(e) => {
                error!("Error leyendo entrada estándar: {}", e);
                ::std::process::exit(1);  
            }
        }
    } else {
        file_path = args[1].to_string();
    }

    let fixed_path = file_path.trim().replace('\\', "/");

    let path = Path::new(&fixed_path);
    if path.exists() {
        info!("Archivo recibido: {}", path.display());
    } else {
        error!("El archivo no existe: {}", path.display());
        ::std::process::exit(1);  
    }
    Some(fixed_path.to_string())
}

fn process(tera: Tera, path: String) -> Tera {


    let price_tables = get_menu_components(&tera, path);

    let rendered_menu = render_menu(&tera, price_tables);

    
    let html_output_path = save_rendered_html(&rendered_menu);
    save_rendered_pdf(&html_output_path);
    tera
}

fn save_rendered_pdf(tmp_html: &std::path::PathBuf) {
    let output_dir = Path::new("dist");
    if let Err(e) = std::fs::create_dir_all(output_dir) {
        error!("Error creando dist: {}", e);
        ::std::process::exit(1);        
    }

    let timestamp = Local::now().format("%Y%m%d-%H%M").to_string();
    let pdf_path = output_dir.join(format!("menu-{}.pdf", timestamp));

    // Preferir google-chrome, fallback a chromium/ chromium-browser
    let chrome_candidates = ["google-chrome-stable", "google-chrome", "chromium", "chromium-browser"];
    let chrome_bin = chrome_candidates.iter()
        .find(|b| which::which(b).is_ok())
        .map(|s| s.to_string())
        .unwrap_or_else(|| {
            error!("No se encontró Chromium/Chrome en PATH. Instala google-chrome o chromium.");
            String::new()
        });

    if chrome_bin.is_empty() {
        ::std::process::exit(1);  
    }

    let html_source_file = format!("./{}", tmp_html.display());
    let status = Command::new(&chrome_bin)
        .arg("--headless")
        .arg("--no-sandbox")                 // necesario en muchos contenedores
        .arg("--disable-gpu")
        .arg("--disable-dev-shm-usage")      
        .arg("--enable-local-file-access") 
        .arg("--force-device-scale-factor=1")
        .arg("--disable-translate")
        .arg("--print-backgrounds")
        .arg(format!("--print-to-pdf={}", pdf_path.display()))
        .arg(html_source_file)
        .stderr(std::process::Stdio::inherit())
        .status();

    match status {
        Ok(s) if s.success() => info!("PDF generado: {}", pdf_path.display()),
        Ok(s) => info!("Chrome finalizó con código: {}", s),
        Err(e) => error!("No se pudo ejecutar Chrome: {}", e),
    }
}

fn get_menu_components(tera: &Tera, path: String) -> Vec<String> {
    info!("Abriendo workbook: {}", path);
    let mut workbook: Xlsx<_> = match open_workbook(path) {
        Ok(wb) => wb,
        Err(e) => {
            error!("Error opening workbook: {}", e);
            ::std::process::exit(1);
        }
    };
    let sheet_names = workbook.sheet_names().to_owned();
    let mut components = Vec::new();

    let product_price_table = vec!["product","price"];

    let offert_price_card = vec!["price", "description"];
    
    for name in sheet_names {
        let range = extract_data_from_sheet(&mut workbook, &name);

        let headers = match range.headers() {
            Some(h) => h,
            None => {
                error!("No headers found in sheet: {}", name);
                ::std::process::exit(1);
            }
        };
        if headers[0].eq(product_price_table[0]) && headers[1].eq(product_price_table[1]) {
            info!("Building table for sheet: {}", name);
            let rendered_table = build_table(tera, &range, &name);
            components.push(rendered_table);       
        } else if headers[0].eq(offert_price_card[0]) && headers[1].eq(offert_price_card[1]) {
            info!("Building offert card for sheet: {}", name);
            let rendered_offert_card = build_offert_card(tera, &range, &name);
            components.push(rendered_offert_card);
        }

    }
    components
}



fn extract_data_from_sheet(workbook: &mut Xlsx<std::io::BufReader<fs::File>>, name: &String) -> Range<Data> {
    let range = match workbook.worksheet_range(name) {
        Ok(r) => r,
        Err(e) => {
            error!("Error reading range: {}", e);
            ::std::process::exit(1);
        }
    };
    range
}
fn build_offert_card(tera: &Tera, data: &Range<Data>, sheet_name: &str) -> String {

    let first_data_row = data
        .rows()
        .skip(1) // saltar la fila de cabecera
        .find(|row| row.iter().any(|cell| !cell.to_string().trim().is_empty()));

    let (price, description) = match first_data_row {
        Some(row) => {
            let price_raw = row.get(0).map(|c| c.to_string()).unwrap_or_default();
            let description = row.get(1).map(|c| c.to_string()).unwrap_or_default();

            // formatear precio si es numérico
            let price = match price_raw.parse::<f64>() {
                Ok(v) => format!("{:.2}€", v),
                Err(_) => price_raw,
            };
            (price, description)
        }
        None => {
            warn!("No data row found in sheet: {}", sheet_name);
            ("".to_string(), "".to_string())
        }
    };
    
    let mut context = Context::new();
        context.insert("name", sheet_name);
        context.insert("price", &price);
        context.insert("description", &description);

    let rendered_offert_card = render_offert_card(tera, &context);
    rendered_offert_card
}



fn render_offert_card(tera: &Tera, context: &Context) -> String {

    // Renderizar la plantilla de tabla a string
    let rendered_offert_card = match tera.render("offert_card_template.html", &context) {
        Ok(s) => s,
        Err(error) => {
            error!("Error renderizando price_table_template: {}", error);
            ::std::process::exit(1);
        }
    };
    rendered_offert_card
}
fn build_table(tera: &Tera, data: &Range<Data> , sheet_name: &str) -> String {

    let mut iter = match RangeDeserializerBuilder::new().from_range(&data){
        Ok(it) => it,
        Err(e) => {
            error!("Error creating deserializer: {}", e);
            ::std::process::exit(1);
        }
    };

    let mut items: Vec<Item> = Vec::new();
    for result in iter.by_ref() {
        let (label, value): (String, String) = match result {
            Ok(result_ok) => result_ok,
            Err(error) => {
                error!("Error reading range: {}", error);
                ::std::process::exit(1);
            }
        };
        items.push(Item {
            name: label,
            price: format!("{:.2}€", value.parse::<f64>().unwrap_or(0.0)),
        });
    }
    let table_context = get_table_from_sheet(sheet_name, &items);

    let rendered_table = render_product_table(tera, &table_context);
    rendered_table
}

fn get_table_from_sheet(table_title: &str, items: &Vec<Item>) -> Context {    
    // Contexto para la plantilla de la tabla
    let mut table_context = Context::new();
    table_context.insert("type", table_title);
    table_context.insert("items", &items);
    table_context
}

fn save_rendered_html(rendered_menu: &String) -> std::path::PathBuf {
    let output_dir = Path::new("dist");
    if !output_dir.exists() {
        if let Err(error) = fs::create_dir_all(output_dir) {
            error!("Error creando directorio de salida: {}", error);
            ::std::process::exit(1);
        }
    }
    
    // Escribir resultado en un archivo (o imprime por stdout si prefieres)
    let output_path = output_dir.join("menu_output.html");
    if let Err(error) = fs::write(&output_path, rendered_menu) {
        error!("Error escribiendo output: {}", error);
    } else {
        info!("menu_output.html generado en dist/menu_output.html");
    }
    output_path
}

fn render_menu(tera: &Tera, price_tables: Vec<String>) -> String {
    // Contexto para el menú que recibe las tablas ya renderizadas
    let mut menu_context = Context::new();
    menu_context.insert("price_tables", &price_tables);

    // Renderizar la plantilla del menú
    let rendered_menu = match tera.render("menu_template.html", &menu_context) {
        Ok(rendered_html) => rendered_html,
        Err(error) => {
            error!("Error renderizando menu_template: {}", error);
            ::std::process::exit(1);
        }
    };
    rendered_menu
}

fn render_product_table(tera: &Tera, table_context: &Context) -> String {   


    // Renderizar la plantilla de tabla a string
    let rendered_table = match tera.render("price_table_template.html", &table_context) {
        Ok(s) => s,
        Err(error) => {
            error!("Error renderizando price_table_template: {}", error);
            ::std::process::exit(1);
        }
    };
    rendered_table
}
