use backfat::container_objects::lines::draw_rectangle;
use backfat::container::placement_info::PlacementInfo;
use backfat::font::font_info::FontInfo;
use backfat::container_objects::text_box::{TextBox, TextAlignment};
use backfat::container_objects::list_box::{TypeOfItem, ListBox, RowData, RowDataTypes, ListBoxBorder};
use backfat::container::rectangle::Border;
use std::cell::RefCell;
use backfat::font::font_sizes::{Font, create_font_recource_id};
use backfat::container::manager::Manager;
use lopdf::{Stream, Object};
use lopdf::content::{Content, Operation};
use lopdf::dictionary;
use std::io::{Write};
use backfat::container::container_trait::DrawInfoReq;

pub struct PdfDrawInfo {
    pub pdf: Vec<Vec<Operation>>,
}
impl DrawInfoReq for PdfDrawInfo {
    fn increment_page_buffer(&mut self, page_number: usize) {
        if page_number >= self.page_array_size() {
            self.pdf.resize(page_number + 1, Vec::new());
        }
    }

    fn page_array_size(&self) -> usize {
        self.pdf.len()
    }

    fn insert_into_page(&mut self, page_num: usize, operation: Operation) {
        self.pdf[page_num].push(operation);
    }
}

pub fn write_rows_to_pdf_container<W:Write>(
    customer: &String,
    _: usize,
    times_qualified: i64,
    data: Vec<Vec<Vec<String>>>,
    save_to: &mut W
) -> Result<(), Box<dyn std::error::Error>> {

    let dpi = 72.0;
    let mut pdf_draw = PdfDrawInfo{ pdf: vec![]};

    let mut pdf_manager = Manager::new(11.0,8.5, dpi, 0.25, 0.25);

    let mut promo_for = TextBox::new(format!("Promotion for: {}", customer), FontInfo::new(10.0, Font::Helvetica), Some(TextAlignment::LeftBottom),None,None, None);
    let mut times_qual = TextBox::new(format!("Times Qualified: {}", times_qualified), FontInfo::new(10.0, Font::Helvetica), Some(TextAlignment::LeftBottom),None,None, None);
    let mut space = TextBox::new("", FontInfo::new(10.0, Font::Helvetica), Some(TextAlignment::LeftBottom),None,None, None);

    let mut placement_handle = pdf_manager.get_placement_handle(2..50, false);
    placement_handle.set_pixel_height(0.27 * dpi);
    let borders: &Option<RefCell<Vec<Border>>> = &Some(RefCell::new(Vec::new()));
    placement_handle.draw( &mut promo_for, &mut pdf_draw, &borders );
    placement_handle = pdf_manager.get_placement_handle(2..20 , false);
    placement_handle.set_pixel_height(0.27 * dpi);
    placement_handle.draw( &mut times_qual, &mut pdf_draw, &borders );
    placement_handle = pdf_manager.get_placement_handle(2..20, false );
    placement_handle.set_pixel_height(0.27 * dpi);
    placement_handle.draw(  &mut space, &mut pdf_draw, &borders );


    let col_size: Vec<usize> = vec![10,20,12,4,11,25,9];
    let header = vec![
        "Ship Date",
        "Customer Name",
        "Order Number",
        "Qty",
        "Part Number",
        "Part Number Description",
        "Sale Price",
    ].into_iter().map(|x| x.to_owned()).collect::<Vec<String>>();


    let mut ends_on_row = 0;
    let mut sec_total_qty: Vec<Option<f64>> = Vec::new();
    for cur_db_rows_index in 0..data.len() {
        let mut row_qty_sum = 0.0;

        for row in &data[cur_db_rows_index] {
            let valid_qty = row[3].parse::<f64>()?;
            row_qty_sum = row_qty_sum + valid_qty;
        }
        if row_qty_sum < 0.0001 {
            sec_total_qty.push(None);
        } else {
            ends_on_row = cur_db_rows_index;
            sec_total_qty.push(Some(row_qty_sum));
        }

    }

    let mut total_qty = 0.0;
    for cur_db_rows_index in 0..data.len() {

        if sec_total_qty[cur_db_rows_index].is_none() {
            continue;
        }

        let mut placement_handle = pdf_manager.get_placement_handle(2..col_size.clone().into_iter().sum::<usize>() + 2, false );

        let data_column_alignment = vec![TextAlignment::LeftJustifyBottom(0.05),TextAlignment::LeftJustifyBottom(0.05),TextAlignment::LeftJustifyBottom(0.05),TextAlignment::RightJustifyBottom(0.05),TextAlignment::LeftJustifyBottom(0.05),TextAlignment::LeftBottom, TextAlignment::RightJustifyBottom(0.05)];
        let trans_data = data[cur_db_rows_index].clone().into_iter().map(|x| RowData::new(x,RowDataTypes::default())).collect::<Vec<RowData>>();
        let dta = RowData::new(header.clone(),RowDataTypes::default());
        let trans_header = Some(&dta);
        let mut list_box = ListBox::new(&trans_data, col_size.clone(), trans_header, &mut pdf_manager, FontInfo::new(10.0,Font::Helvetica), FontInfo::new(12.0, Font::Helvetica), ListBoxBorder::All(1.4,1.4), None);

        list_box.set_item_column_alignments(data_column_alignment);
        list_box.set_header_column_alignments(vec![TextAlignment::LeftJustifyCenter(0.05);header.len()]);
        list_box.set_border_color((0.0,0.0,0.0));
        list_box.header_has_border(false);
        let row_data_types = vec![TypeOfItem::String,TypeOfItem::String,TypeOfItem::String,TypeOfItem::Number(2),TypeOfItem::String,TypeOfItem::String,TypeOfItem::Currency(2)];
        list_box.set_row_types(row_data_types);

        placement_handle.draw( &mut list_box, &mut pdf_draw, &borders);

        total_qty = sec_total_qty[cur_db_rows_index].unwrap();
        let mut qty_total = TextBox::new(format!("Total Quantity: {}", total_qty), FontInfo::new(14.0, Font::Helvetica), Some(TextAlignment::RightJustifyCenter(0.05)),None,None, None);

        let qty_column_range_start = col_size[0..3].into_iter().sum::<usize>() + 2;

        placement_handle = pdf_manager.get_placement_handle(2..qty_column_range_start + col_size[3], false );
        placement_handle.set_pixel_height(0.27 * dpi);
        placement_handle.draw( &mut qty_total, &mut pdf_draw, &borders );

        //is not last row.
        if cur_db_rows_index < ends_on_row {
            let mut space = TextBox::new("", FontInfo::new(10.0, Font::Helvetica), Some(TextAlignment::LeftBottom), None, None, None);

            let mut placement_handle = pdf_manager.get_placement_handle(0..50, false);
            placement_handle.set_pixel_height(0.27 * dpi);
            placement_handle.draw(&mut space, &mut pdf_draw, &borders);

        }

    }

    for group_rec in pdf_manager.get_groups() {
        for page_index in 0..group_rec.1.len() {

            let mut pl: PlacementInfo = PlacementInfo::default();
            pl.page_number = page_index;
            pl.rec = group_rec.1[page_index].clone();
            pl.page_size_info = pdf_manager.get_page_info();

            draw_rectangle(&mut pdf_draw,
                           &pl,
                           5.0,
                           (1.0,0.0,0.0));
        }
    }

    match borders {
        Some(brd) => {

            let hm = brd.to_owned().into_inner();
            for border in hm.into_iter() {

                draw_rectangle(&mut pdf_draw,
                               &border.rec,
                               border.pixel_size,
                               border.color);
            };
        },
        None => {

        }
    }

    let mut doc = lopdf::Document::with_version("1.5");
    let pages_id = doc.new_object_id();

    let mut v:Vec<lopdf::Object> = Vec::new();
    let resources_id = create_font_recource_id(&mut doc);

    for page in 0..pdf_manager.get_page_cnt() + 1 {
        let content = Content {
            operations: pdf_draw.pdf[page].clone()//draw_data_pdf(0.2, 1.0, &mut columns)
        };
        let content_id = doc.add_object(Stream::new(dictionary! {}, content.encode().unwrap()));
        let page_id = doc.add_object(dictionary! {
            "Type" => "Page",
            "Parent" => pages_id,
            "Contents" => content_id,
            });
        v.push( page_id.into() )
    };
    let page_count = v.len() as i32;
    let pages = dictionary! {
		"Type" => "Pages",
		"Kids" => v,
		"Count" => page_count,
		"Resources" => resources_id,
		"MediaBox" => vec![0.into(), 0.into(), (pdf_manager.get_page_pixel_dims().0).into(), (pdf_manager.get_page_pixel_dims().1).into()],
	};
    doc.objects.insert(pages_id, Object::Dictionary(pages));
    let catalog_id = doc.add_object(dictionary! {
		"Type" => "Catalog",
		"Pages" => pages_id,
	});
    doc.trailer.set("Root", catalog_id);
    doc.compress();

    doc.save_to(save_to)?;

    Ok(())
}
