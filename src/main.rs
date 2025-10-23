use serde::Serialize;
use std::fs;
use std::path::Path;
use tera::Context;
use tera::Tera;
use calamine::{open_workbook, Xlsx, Reader, RangeDeserializerBuilder};

#[derive(Serialize)]
struct Item {
    name: String,
    price: String,
}

fn main() {
    let path = format!("{}/dist/menu_test.xlsx", env!("CARGO_MANIFEST_DIR"));
    match Tera::new("template/*.html") {
        Ok(tera) => process(tera, path),
        Err(error) => {
            println!("Parsing error(s): {}", error);
            ::std::process::exit(1);
        }
    };
}

fn process(tera: Tera, path: String) -> Tera {


    let price_tables = get_menu_components(&tera, path);

    let rendered_menu = render_menu(&tera, price_tables);


    save_rendered_html(rendered_menu);
    tera
}

fn get_menu_components(tera: &Tera, path: String) -> Vec<String> {
    let mut workbook: Xlsx<_> = match open_workbook(path) {
        Ok(wb) => wb,
        Err(e) => {
            eprintln!("Error opening workbook: {}", e);
            ::std::process::exit(1);
        }
    };
    let sheet_names = workbook.sheet_names().to_owned();
    let mut price_tables = Vec::new();
    
    for name in sheet_names {
        let rendered_table = build_table(tera, &mut workbook, &name);
        price_tables.push(rendered_table);
    }
    price_tables
}

fn build_table(tera: &Tera, workbook: &mut Xlsx<std::io::BufReader<fs::File>>, sheet_name: &str) -> String {
    let range = match workbook.worksheet_range(sheet_name) {
        Ok(r) => r,
        Err(e) => {
            eprintln!("Error reading range: {}", e);
            ::std::process::exit(1);
        }
    };


    let mut iter = match RangeDeserializerBuilder::new().from_range(&range){
        Ok(it) => it,
        Err(e) => {
            eprintln!("Error creating deserializer: {}", e);
            ::std::process::exit(1);
        }
    };

    let mut items: Vec<Item> = Vec::new();
    for result in iter.by_ref() {
        let (label, value): (String, String) = match result {
            Ok(result_ok) => result_ok,
            Err(error) => {
                eprintln!("Error reading range: {}", error);
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

fn save_rendered_html(rendered_menu: String) {
    let output_dir = Path::new("dist");
    if !output_dir.exists() {
        if let Err(error) = fs::create_dir_all(output_dir) {
            eprintln!("Error creando directorio de salida: {}", error);
            ::std::process::exit(1);
        }
    }
    
    // Escribir resultado en un archivo (o imprime por stdout si prefieres)
    if let Err(error) = fs::write(output_dir.join("menu_output.html"), rendered_menu) {
        eprintln!("Error escribiendo output: {}", error);
    } else {
        println!("menu_output.html generado en dist/menu_output.html");
    }
}

fn render_menu(tera: &Tera, price_tables: Vec<String>) -> String {
    // Contexto para el menú que recibe las tablas ya renderizadas
    let mut menu_context = Context::new();
    menu_context.insert("price_tables", &price_tables);

    // Renderizar la plantilla del menú
    let rendered_menu = match tera.render("menu_template.html", &menu_context) {
        Ok(rendered_html) => rendered_html,
        Err(error) => {
            eprintln!("Error renderizando menu_template: {}", error);
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
            eprintln!("Error renderizando price_table_template: {}", error);
            ::std::process::exit(1);
        }
    };
    rendered_table
}
