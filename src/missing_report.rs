use std::collections::HashMap;
use std::io::prelude::*;
use lopdf::{Object, Stream};
use lopdf::content::{Content};
use lopdf::dictionary;
use zip::write::{FileOptions, ZipWriter};

fn get_amount_missing_for_next_promo(qty_needed: i64, qty_claimed: i64) -> i64 {
    if qty_claimed < qty_needed {
        qty_needed - qty_claimed
    } else if qty_claimed == qty_needed {
        qty_needed
    } else {
        // 6
        let promo_count_ceiling = qty_claimed / qty_needed + 1;
        let remaining_needed = promo_count_ceiling * qty_needed - qty_claimed;
        remaining_needed
    }
}

struct MissingPartNumber {
    pub missing_part_numbers: Vec<String>,
    pub amount_needed: i64,
}
struct NeededSections {
    pub missing_part_numbers: Vec<(AndOrType, Vec<MissingPartNumber>)>,
}

fn generate_missing_report_for_section(promo_section: &PromoSection) -> Vec<NeededSections> {
    let mut rv: Vec<NeededSections> = Vec::new();
    for stl_nd_sec_ind in 0..promo_section.promo_parts_still_needed.len() {
        let part_index: usize = promo_section.promo_parts_still_needed[stl_nd_sec_ind];
        let promo_section_part = &promo_section.part[part_index];
        let mut sec = NeededSections {
            missing_part_numbers: Vec::new(),
        };
        let mut missing_pn: Vec<MissingPartNumber> = Vec::new();
        for tpn_index in 0..promo_section_part.type_prods_for_next_promo_needed.len() {
            let total_qty = promo_section_part.type_prod
                [promo_section_part.type_prods_for_next_promo_needed[tpn_index] as usize]
                .total_qty;
            let qty_needed = promo_section_part.type_prod
                [promo_section_part.type_prods_for_next_promo_needed[tpn_index] as usize]
                .qty_needed as i64;
            let rem = get_amount_missing_for_next_promo(qty_needed, total_qty);
            let missing = MissingPartNumber {
                missing_part_numbers: promo_section_part.type_prod
                    [promo_section_part.type_prods_for_next_promo_needed[tpn_index] as usize]
                    .part_numbers
                    .clone(),
                amount_needed: rem,
            };
            missing_pn.push(missing);
        }
        sec.missing_part_numbers
            .push((promo_section_part.part_type.clone(), missing_pn));
        //
        rv.push(sec);
    }
    rv
}
fn display_vec_of_strings_as_csv<W: Write>(
    input: &Vec<String>,
    write_to: &mut W,
) -> Result<(), std::io::Error> {
    for val_idx in 0..input.len() {
        if val_idx < input.len() - 1 {
            let t = input[val_idx].clone() + " ,";
            write_to.write(t.as_bytes())?;
        } else {
            write_to.write(input[val_idx].as_bytes())?;
        }
    }
    Ok(())
}
fn write_missing_report<W: Write>(
    missing_report: &Vec<NeededSections>,
    write_to: &mut W,
) -> Result<(), std::io::Error> {
    for sec_index in 0..missing_report.len() {
        let sec = &missing_report[sec_index];
        for missing_index in 0..sec.missing_part_numbers.len() {
            let (join_type, items) = &sec.missing_part_numbers[missing_index];
            for item_idx in 0..items.len() {
                write_to.write(format!("{} of a ( ", items[item_idx].amount_needed).as_bytes())?;

                display_vec_of_strings_as_csv(&items[item_idx].missing_part_numbers, write_to)?;
                write_to.write(format!(" )").as_bytes())?;

                if item_idx < items.len() - 1 {
                    if join_type == &AndOrType::Or {
                        write_to.write(format!(" {} ", join_type).as_bytes())?;
                    } else {
                        write_to.write(format!("\r\n").as_bytes())?;
                    }
                } else {
                    write_to.write(format!("\r\n").as_bytes())?;
                }
            }
        }
    }
    Ok(())
}

pub fn display_missing_report<W: Write>(
    hsh: &mut HashMap<String, Promotion>,
    write_to: &mut W,
) -> Result<(), std::io::Error> {
    let mut cust_names: Vec<_> = hsh.iter().map(|x| x.0).collect();
    cust_names.sort();
    for name in cust_names {
        write_to.write(format!("For Customer: {}\r\n", name).as_bytes())?;

        for sec_id in 0..hsh[name].promo_sections.len() {
            write_to.write(
                format!(
                    "Qualified {} times for Promo {}\r\n",
                    &hsh[name].promo_sections[sec_id].times_section_qualified,
                    sec_id + 1
                )
                .as_bytes(),
            )?;

            if hsh[name].promo_sections[sec_id].times_section_qualified == 0 {
                write_to.write(format!("To get the promo you need to purchase:\r\n").as_bytes())?;
            } else {
                write_to.write(format!("To get another you need to purchase:\r\n").as_bytes())?;
            }
            let missing_section_data =
                generate_missing_report_for_section(&hsh[name].promo_sections[sec_id]);
            write_missing_report(&missing_section_data, write_to)?;
        }
        write_to.write(format!("\r\n").as_bytes())?;
    }
    Ok(())
}
struct PdfDox {
    manager: Manager,
}

impl PdfDox {
    pub fn new(width_inch: f64, height_inch: f64, dpi: f64, margin_top:f64, margin_bot:f64) -> Self {
        Self {
            manager: Manager::new(width_inch, height_inch, dpi, margin_top, margin_bot),
        }
    }
}
pub fn write_missing_report_to_pdf<W: Write>(
    hsh: &HashMap<String, Promotion>,
    write_to: &mut W,
) -> Result<(), std::io::Error> {
    let mut cust_names: Vec<_> = hsh.iter().map(|x| x.0).collect();
    cust_names.sort();
    let borders: Option<RefCell<Vec<Border>>> = Some(RefCell::new(Vec::new()));
    let mut pdf_draw = PdfDrawInfo{ pdf: vec![] };
    let mut dox = PdfDox::new( 8.5, 11.0, 72.0, 0.25,0.25 );

    let mut should_new_page = false;

    for name in cust_names {

        let mut txt = TextBox::new(format!("For Customer: {}\r\n", name), FontInfo::new(16.0, Font::Helvetica), Some(TextAlignment::LeftBottom), None, None, None);
        let mut placement_handle = dox.manager.get_placement_handle(1..99, should_new_page);
        should_new_page = true;
        placement_handle.set_pixel_height(0.25 * 72.0);
        placement_handle.draw(&mut txt, &mut pdf_draw, &borders);
        let mut group = 0;
        for sec_id in 0..hsh[name].promo_sections.len() {
            let mut txt = TextBox::new(
                format!(
                    "Qualified {} times for Promo {}\r\n",
                    &hsh[name].promo_sections[sec_id].times_section_qualified,
                    sec_id + 1
                ), FontInfo::new(14.0, Font::Helvetica), Some(TextAlignment::LeftCenter), None, None, None);
            placement_handle = dox.manager.get_placement_handle(1..99, false);
            placement_handle.set_pixel_height(0.25 * 72.0);
            placement_handle.draw(&mut txt, &mut pdf_draw, &borders);

            txt = TextBox::new(
                "", FontInfo::new(14.0, Font::Helvetica), Some(TextAlignment::LeftBottom), None, None, None);
            placement_handle = dox.manager.get_placement_handle(1..99, false);
            placement_handle.set_pixel_height(0.25 * 72.0);
            placement_handle.draw(&mut txt, &mut pdf_draw, &borders);


            if hsh[name].promo_sections[sec_id].times_section_qualified == 0 {
                txt = TextBox::new(
                    format!("To get the promo"), FontInfo::new(12.0, Font::Helvetica), Some(TextAlignment::LeftCenter), None, None, None);
                placement_handle = dox.manager.get_placement_handle(1..99, false);
                placement_handle.set_pixel_height(0.25 * 72.0);
                placement_handle.draw(&mut txt, &mut pdf_draw, &borders);
            } else {
                txt = TextBox::new(
                    format!("To get another promo"), FontInfo::new(12.0, Font::Helvetica), Some(TextAlignment::LeftCenter), None, None, None);
                placement_handle = dox.manager.get_placement_handle(2..80, false);
                placement_handle.set_pixel_height(0.25 * 72.0);
                placement_handle.draw(&mut txt, &mut pdf_draw, &borders);
            }
            let missing_section_data = generate_missing_report_for_section(&hsh[name].promo_sections[sec_id]);

            write_missing_report_to_pdf_new(1..99,&missing_section_data, &mut pdf_draw, &borders, &mut dox, Some(group))?;

            group = group + 1;
            txt = TextBox::new(
                "", FontInfo::new(10.0, Font::Helvetica), Some(TextAlignment::CenterCenter), None, None, None);
            placement_handle = dox.manager.get_placement_handle(8..92, false);
            placement_handle.set_pixel_height(0.30 * 72.0);
            placement_handle.draw(&mut txt, &mut pdf_draw, &borders);


        }
        txt = TextBox::new(
            "", FontInfo::new(10.0, Font::Helvetica), Some(TextAlignment::CenterCenter), None, None, None);
        placement_handle = dox.manager.get_placement_handle(8..92, false);
        placement_handle.set_pixel_height(0.25 * 72.0);
        placement_handle.draw(&mut txt, &mut pdf_draw, &borders);

    }


    match borders {
        Some(brd) => {

            for border in brd.into_inner().into_iter() {

                draw_rectangle(&mut pdf_draw,
                               &border.rec,
                               border.pixel_size,
                               border.color);
            };
        },
        None => {

        }
    }


    for group_rec in dox.manager.get_groups() {
        for page_index in 0..group_rec.1.len() {

            let mut pl: PlacementInfo = PlacementInfo::default();
            pl.page_number = page_index;
            pl.rec = group_rec.1[page_index].clone();
            pl.page_size_info = dox.manager.get_page_info();
            let col = if group_rec.0 == 0 {
                (1.0,0.0,0.0)
            } else if group_rec.0 == 1 {
                (0.0,1.0,0.0)
            } else if group_rec.0 == 2 {
                (1.0,1.0,0.0)
            } else if group_rec.0 == 3 {
                (1.0,0.0,1.0)
            } else {
                (0.0,0.0,0.0)
            };
            let grw_amt = 6.0;
            pl.rec.x = pl.rec.x - grw_amt;
            pl.rec.width = pl.rec.width + 2.0 * grw_amt;
            pl.rec.y = pl.rec.y - grw_amt;
            pl.rec.height = pl.rec.height + 2.0 * grw_amt;
            draw_rectangle(&mut pdf_draw,
                           &pl,
                           3.0,
                           col);
        }
    }


    let mut doc = lopdf::Document::with_version("1.5");
    let pages_id = doc.new_object_id();

    let resources_id = create_font_recource_id(&mut doc);
    let mut v:Vec<lopdf::Object> = Vec::new();

    for page in 0..dox.manager.get_page_cnt() + 1 {
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
		"MediaBox" => vec![0.into(), 0.into(), (dox.manager.get_page_pixel_dims().0).into(), (dox.manager.get_page_pixel_dims().1).into()],
	};
    doc.objects.insert(pages_id, Object::Dictionary(pages));
    let catalog_id = doc.add_object(dictionary! {
		"Type" => "Catalog",
		"Pages" => pages_id,
	});
    doc.trailer.set("Root", catalog_id);
    doc.compress();
    doc.save_to(write_to).unwrap();

    Ok(())
}
pub fn write_missing_report_to_pdf_per_customer<W: Write>(
    hsh: &HashMap<String, Promotion>,
    customer: &String,
    write_to: &mut W,
) -> Result<(), std::io::Error> {
    let mut cust_names: Vec<_> = hsh.iter().map(|x| x.0).collect();
    cust_names.sort();
    let borders: Option<RefCell<Vec<Border>>> = Some(RefCell::new(Vec::new()));
    let mut pdf_draw = PdfDrawInfo{ pdf: vec![] };
    let mut dox = PdfDox::new( 8.5, 11.0, 72.0, 0.25,0.25 );

    let mut txt = TextBox::new(format!("For Customer: {}\r\n", customer), FontInfo::new(16.0, Font::Helvetica), Some(TextAlignment::LeftBottom), None, None, None);
    let mut placement_handle = dox.manager.get_placement_handle(1..99, false);
    placement_handle.set_pixel_height(0.25 * 72.0);
    placement_handle.draw(&mut txt, &mut pdf_draw, &borders);
    let mut group = 0;
    for sec_id in 0..hsh[customer].promo_sections.len() {
        let mut txt = TextBox::new(
            format!(
                "Qualified {} times for Promo {}\r\n",
                &hsh[customer].promo_sections[sec_id].times_section_qualified,
                sec_id + 1
            ), FontInfo::new(14.0, Font::Helvetica), Some(TextAlignment::LeftCenter), None, None, None);
        placement_handle = dox.manager.get_placement_handle(1..99, false);
        placement_handle.set_pixel_height(0.25 * 72.0);
        placement_handle.draw(&mut txt, &mut pdf_draw, &borders);

        txt = TextBox::new(
            "", FontInfo::new(14.0, Font::Helvetica), Some(TextAlignment::LeftBottom), None, None, None);
        placement_handle = dox.manager.get_placement_handle(1..99, false);
        placement_handle.set_pixel_height(0.25 * 72.0);
        placement_handle.draw(&mut txt, &mut pdf_draw, &borders);

        if hsh[customer].promo_sections[sec_id].times_section_qualified == 0 {
            txt = TextBox::new(
                format!("To get the promo"), FontInfo::new(12.0, Font::Helvetica), Some(TextAlignment::LeftCenter), None, None, None);
            placement_handle = dox.manager.get_placement_handle(1..99, false);
            placement_handle.set_pixel_height(0.25 * 72.0);
            placement_handle.draw(&mut txt, &mut pdf_draw, &borders);
        } else {
            txt = TextBox::new(
                format!("To get another promo"), FontInfo::new(12.0, Font::Helvetica), Some(TextAlignment::LeftCenter), None, None, None);
            placement_handle = dox.manager.get_placement_handle(2..80, false);
            placement_handle.set_pixel_height(0.25 * 72.0);
            placement_handle.draw(&mut txt, &mut pdf_draw, &borders);
        }
        let missing_section_data = generate_missing_report_for_section(&hsh[customer].promo_sections[sec_id]);

        write_missing_report_to_pdf_new(1..99,&missing_section_data, &mut pdf_draw, &borders, &mut dox, Some(group))?;

        group = group + 1;
        txt = TextBox::new(
            "", FontInfo::new(10.0, Font::Helvetica), Some(TextAlignment::CenterCenter), None, None, None);
        placement_handle = dox.manager.get_placement_handle(8..92, false);
        placement_handle.set_pixel_height(0.30 * 72.0);
        placement_handle.draw(&mut txt, &mut pdf_draw, &borders);
    }

    match borders {
        Some(brd) => {

            for border in brd.into_inner().into_iter() {

                draw_rectangle(&mut pdf_draw,
                               &border.rec,
                               border.pixel_size,
                               border.color);
            };
        },
        None => {

        }
    }

    for group_rec in dox.manager.get_groups() {
        for page_index in 0..group_rec.1.len() {

            let mut pl: PlacementInfo = PlacementInfo::default();
            pl.page_number = page_index;
            pl.rec = group_rec.1[page_index].clone();
            pl.page_size_info = dox.manager.get_page_info();
            let mut col = if group_rec.0 == 0 {
                (1.0,0.0,0.0)
            } else if group_rec.0 == 1 {
                (0.0,1.0,0.0)
            } else if group_rec.0 == 2 {
                (1.0,1.0,0.0)
            } else if group_rec.0 == 3 {
                (1.0,0.0,1.0)
            } else {
                (0.0,0.0,0.0)
            };
            col = (0.0,0.0,0.0);
            let grw_amt = 6.0;
            pl.rec.x = pl.rec.x - grw_amt;
            pl.rec.width = pl.rec.width + 2.0 * grw_amt;
            pl.rec.y = pl.rec.y - grw_amt;
            pl.rec.height = pl.rec.height + 2.0 * grw_amt;
            draw_rectangle(&mut pdf_draw,
                           &pl,
                           3.0,
                           col);
        }
    }


    let mut doc = lopdf::Document::with_version("1.5");
    let pages_id = doc.new_object_id();

    let resources_id = create_font_recource_id(&mut doc);
    let mut v:Vec<lopdf::Object> = Vec::new();

    for page in 0..dox.manager.get_page_cnt() + 1 {
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
		"MediaBox" => vec![0.into(), 0.into(), (dox.manager.get_page_pixel_dims().0).into(), (dox.manager.get_page_pixel_dims().1).into()],
	};
    doc.objects.insert(pages_id, Object::Dictionary(pages));
    let catalog_id = doc.add_object(dictionary! {
		"Type" => "Catalog",
		"Pages" => pages_id,
	});
    doc.trailer.set("Root", catalog_id);
    doc.compress();
    doc.save_to(write_to).unwrap();

    Ok(())
}

fn write_missing_report_to_pdf_new( placement_range: Range<usize>,
    missing_report: &Vec<NeededSections>,
    pdf_draw: &mut PdfDrawInfo, borders: &Option<RefCell<Vec<Border>>>, dox: &mut PdfDox, group: Option<usize>,
) -> Result<(), std::io::Error> {
    let col_size: Vec<usize> = vec![10, 10, 10, 10, 10];//vec![9, 10, 11, 12];
    let col_size_len = col_size.len();
    let col_sum = col_size.clone().into_iter().sum::<usize>();
    let width_of_range = placement_range.end - placement_range.start;
    let half_range = Range{start: width_of_range / 2 - col_sum / 2, end: width_of_range / 2 + col_sum / 2};

    for sec_index in 0..missing_report.len() {
        let sec = &missing_report[sec_index];

        let mut data: Vec<RowData> = Vec::new();
        let mut row_on = 0;

        let purchase_color = (0.3,0.3,0.9);
        for missing_index in 0..sec.missing_part_numbers.len() {

            let (join_type, items) = &sec.missing_part_numbers[missing_index];
            for item_idx in 0..items.len() {
                if item_idx == 0 {
                    let str = format!("Purchase {} more", items[item_idx].amount_needed);
                    data.push(RowData::new(vec![str], RowDataTypes::SingleWithColor(purchase_color, TextAlignment::CenterCenter)));
                }
                else if item_idx < items.len() {
                    match join_type {
                        AndOrType::Or => {
                            let str = format!("{} purchase {} more", join_type, items[item_idx].amount_needed);
                            data.push(RowData::new(vec![str],RowDataTypes::SingleWithColor(purchase_color, TextAlignment::CenterCenter)));
                        }
                        AndOrType::And => {
                            let str = format!("{} purchase {} more", join_type, items[item_idx].amount_needed);
                            data.push(RowData::new(vec![str],RowDataTypes::SingleWithColor(purchase_color, TextAlignment::CenterCenter)));
                        }
                        AndOrType::Any(_) => {
                            let str = format!("Or purchase {} more", items[item_idx].amount_needed);
                            data.push(RowData::new(vec![str],RowDataTypes::SingleWithColor(purchase_color, TextAlignment::CenterCenter)));
                        }
                        AndOrType::None => {
                            let str = format!("{} purchase {} more", join_type, items[item_idx].amount_needed);
                            data.push(RowData::new(vec![str],RowDataTypes::SingleWithColor(purchase_color, TextAlignment::CenterCenter)));
                        }

                    }

                }

                let mut tt = 0;

                let mut acc_data: Vec<String> = Vec::new();
                for m_ind in 0..items[item_idx].missing_part_numbers.len() {
                    let m = &items[item_idx].missing_part_numbers[m_ind];
                    if m_ind % col_size_len == 0 && m_ind != 0 {
                        acc_data.push(m.to_string());
                        data.push(RowData::new(acc_data,RowDataTypes::default()));
                        acc_data = Vec::new();
                        tt = tt + 1;
                        row_on = row_on + 1;
                    } else {
                        acc_data.push(m.to_string());
                    }
                }
                if !acc_data.is_empty() {
                    data.push(RowData::new(acc_data, RowDataTypes::default()));
                }
           //     row_on = row_on + tt;//how_many_rows / col_size_len + 1;


              //  row_on = row_on + 1;

            }
            let mut placement_handle = dox.manager.get_placement_handle(Range { start: half_range.start, end: half_range.end }, false);

            let mut list_box = ListBox::new(&data,
                                            col_size.clone(),
                                            None,
                                            &mut dox.manager,
                                            FontInfo::new(12.0, Font::Helvetica),
                                            FontInfo::new(12.0, Font::Helvetica),
                                            ListBoxBorder::All(2.0,3.0), group
            );

 //           for r in solid_rows {

                //list_box.add_solid_row(r.0,r.1);
   //         }
            list_box.set_row_types(vec![TypeOfItem::String, TypeOfItem::String,TypeOfItem::String, TypeOfItem::String, TypeOfItem::String, TypeOfItem::String]);
            list_box.set_item_column_alignments(vec![TextAlignment::CenterCenter, TextAlignment::CenterCenter,TextAlignment::CenterCenter, TextAlignment::CenterCenter, TextAlignment::CenterCenter, TextAlignment::CenterCenter, ]);
            placement_handle.draw(&mut list_box, pdf_draw, &borders);


            /*
                            txt = TextBox::new(
                                format!("Purchase {} more", items[item_idx].amount_needed),
                                FontInfo::new(12.0, Font::Helvetica), Some(TextAlignment::LeftCenter), None, None);
                            placement_handle = dox.manager.get_placement_handle(Range { start: placement_range.start, end: half_range.start }, false);
                            placement_handle.set_pixel_height(0.25 * 72.0);
                            placement_handle.draw(&mut txt, pdf_draw, &borders);
                            /////////
                            placement_handle = dox.manager.get_placement_handle(half_range.clone(), false);


            //                let data = items[item_idx].missing_part_numbers.clone().into_iter().fold(String::new(), |acc,i| acc + ", " + &i);

                            let mut list_box = ListBox::new(&data,
                                                            col_size.clone(),
                                                            None,
                                                            &mut dox.manager,
                                                            FontInfo::new(12.0, Font::Helvetica),
                                                            FontInfo::new(12.0, Font::Helvetica),
                                                            ListBoxBorder::None,
                            );

                            list_box.set_row_types(vec![TypeOfItem::String, TypeOfItem::String, TypeOfItem::String, TypeOfItem::String, TypeOfItem::String]);
                            list_box.set_item_column_alignments(vec![TextAlignment::CenterCenter, TextAlignment::CenterCenter, TextAlignment::CenterCenter, TextAlignment::CenterCenter, TextAlignment::CenterCenter, ]);
                            placement_handle.draw(&mut list_box, pdf_draw, &borders);

                            let mut placement_handle = dox.manager.get_placement_handle(half_range.clone(), false);
            */

        }

        if sec_index < missing_report.len() - 1 {
       //     data[row_on] = vec!["".to_owned(), "".to_owned(),"And".into(),"".to_owned(),"".to_owned()];
            let mut placement_handle = dox.manager.get_placement_handle(Range { start: half_range.start, end: half_range.end }, false);
            let mut txt = TextBox::new(
                format!("And"),
                FontInfo::new(12.0, Font::Helvetica), Some(TextAlignment::CenterCenter), None, None, group);
          //  placement_handle = dox.manager.get_placement_handle(Range { start: placement_range.start, end: placement_range.end }, false);
            placement_handle.set_pixel_height(0.30 * 72.0);
            placement_handle.draw(&mut txt, pdf_draw, &borders);
        }
        //       write_to.write(format!("\r\n").as_bytes())?;
    } //end sec
    Ok(())
}

use std::cell::RefCell;
use backfat::container::rectangle::{Border};
use backfat::container_objects::text_box::{TextBox, TextAlignment};
use backfat::font::font_info::FontInfo;
use backfat::font::font_sizes::{Font, create_font_recource_id};
use backfat::container::manager::{Manager};
use backfat::container_objects::list_box::{TypeOfItem, ListBoxBorder, ListBox, RowData, RowDataTypes};
use std::ops::Range;
use backfat::container_objects::lines::draw_rectangle;
use backfat::container::placement_info::PlacementInfo;
use promo_input::general::promo_json::{Promotion, PromoSection};
use promo_input::general::and_or::AndOrType;
use promo_input::general::data::load_promo;
use crate::pdf::{write_rows_to_pdf_container, PdfDrawInfo};


pub fn run_missing_reports<W: Write>(
    input_file: &str,
    json_promo_file: &str,
    output_file: Option<&mut W>,
    zip_path: &str,
) -> Result<(), Box<dyn std::error::Error>> {

//    std::fs::create_dir_all(&path)?;
    let zip_file = std::fs::File::create(zip_path.to_owned() ).expect("Couldn't create file");
    let mut zip_file_writer = ZipWriter::new(zip_file);

    let completed_promo = load_promo( input_file, json_promo_file)?;
    if let Some( full_file ) = output_file {
        write_missing_report_to_pdf(  &completed_promo.data, full_file )?;
    }

    for (customer, promo) in &completed_promo.data {
        let mut v = Vec::new();
        write_missing_report_to_pdf_per_customer(  &completed_promo.data, customer,&mut v )?;
        let write_file = format!("Missing_Reports\\{} Missing Report.pdf", customer);
        zip_file_writer.start_file(write_file.to_owned(), FileOptions::default())?;

        zip_file_writer.write(&v)?;
        for section_index in 0..promo.promo_sections.len() {
            let section = &promo.promo_sections[section_index];
            if section.times_section_qualified > 0 {
                let mut parts_ret: Vec<Vec<Vec<String>>> = Vec::new();

                for part in &section.part {
                    for type_prod in &part.type_prod {
                        let mut all_rows2: Vec<Vec<String>> = Vec::new();

                        for row in &type_prod.found_numbers {
                            all_rows2.push(vec![
                                row[completed_promo.ship_date_column_index].value.clone(),
                                row[completed_promo.customer_name_column_index].value.clone(),
                                row[completed_promo.order_number_column_index].value.clone(),
                                row[completed_promo.qty_column_index].value.clone(),
                                row[completed_promo.part_number_column_index].value.clone(),
                                row[completed_promo.part_number_desc_column_index].value.clone(),
                                row[completed_promo.sales_column_index].value.clone(),
                            ]);
                        }
                        all_rows2.sort_by(|x, y| x[0].cmp(&y[0]));
                        parts_ret.push(all_rows2);
                    }
                }

                write_rows_to_pdf_container(
                    customer,
                    section_index,
                    section.times_section_qualified,
                    parts_ret.clone(),
                    &mut v,

                )?;
                let write_file = format!("{}\\Promo#{}.pdf", customer, section_index.to_string());

                zip_file_writer.start_file(write_file.to_owned(), FileOptions::default())?;
                zip_file_writer.write(&v)?;

            }
        }
    }
    zip_file_writer.finish()?;
    Ok(())
}