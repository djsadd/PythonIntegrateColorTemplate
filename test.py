import win32com.client
import os
import time

def replace_text_in_coreldraw(kaz, rus, eng, filenamexcel=None):
    corel = win32com.client.Dispatch("CorelDRAW.Application")
    corel.Visible = True
    try:
        template_path = r"C:\Users\admin\Desktop\Дипломы\ШАБЛОН ДИПЛОМА МАГИСТРАТУРА.cdr"
        document = corel.OpenDocument(template_path)
        
        # # 4. Путь к изображению PNG
        # image_path = r"C:\Users\admin\Дипломки\без фона"
        # image_path = image_path + r"\00016635125.png"
        # print(image_path)
        # import_filter = 0 
        # options = None 
        
        # shape = document.ActiveLayer.Import(image_path, import_filter, options)
        
        # print("PNG с альфа-каналом импортирован в существующий CDR-документ")

        
                
        for page in document.Pages:
            for shape in page.Shapes:
                if shape.type == 6:
                    text = shape.Text.Story.Text
                    
                    if "ФИО НА КАЗ" in text:
                        fio = format_fio(kaz[1])
                        shape.Text.Story.Text = shape.Text.Story.Text.replace("ФИО НА КАЗ", fio)
                        text_width = shape.SizeWidth
                        page_width = page.SizeWidth
                        shape.PositionX = (page_width - text_width) / 2
                        
                    if "«ОП»" in text:
                        op_kz = format_op_name(kaz[2])
                        shape.Text.Story.Text = shape.Text.Story.Text.replace("«ОП»", f"«{op_kz}»")
                        text_width = shape.SizeWidth
                        page_width = page.SizeWidth
                        shape.PositionX = (page_width - text_width) / 2
                        
                    if "степень магистра каз" in text:
                        degree_kaz = kaz[3].lower()
                        shape.Text.Story.Text = shape.Text.Story.Text.replace("степень магистра каз", f"{degree_kaz}")
                        text_width = shape.SizeWidth
                        page_width = page.SizeWidth
                        shape.PositionX = (page_width - text_width) / 2

                    if "№нпрот" in text or "гпк" in text or "дпк" in text:
                        updated_text = text
                        if "№нпрот" in updated_text:
                            updated_text = updated_text.replace("№нпрот", f"№{kaz[4]}")
                        if "гпк" in updated_text:
                            year_kaz = extract_date_time(kaz[5])[2]
                            updated_text = updated_text.replace("гпк", f"{year_kaz}")
                        if "дпк" in updated_text:
                            month_kaz = extract_date_time(kaz[5])[0]
                            updated_text = updated_text.replace("дпк", f"{month_kaz}")

                        if "мпк" in updated_text:
                            month_kaz = extract_date_time(kaz[5])[1]
                            month_kaz = kz_months_locative[month_kaz]
                            updated_text = updated_text.replace("мпк", f"{month_kaz}")
                            
                        shape.Text.Story.Text = updated_text

                    if "рег.н." in text:
                        new_text = text.replace("рег.н.", f"{kaz[7]}")
                        shape.Text.Story.Text = new_text

                    if "дк" in text:
                        day_kaz = extract_date_time(kaz[8])[0]
                        new_text = text.replace("дк", f"{day_kaz}")
                        shape.Text.Story.Text = new_text
                    if "мк" in text:
                        month_kaz = extract_date_time(kaz[8])[1]
                        new_text = text.replace("мк", f"{month_kaz}")
                        shape.Text.Story.Text = new_text
                    if "гк" in text:
                        year_kaz = extract_date_time(kaz[8])[2]
                        new_text = text.replace("гк", f"{year_kaz}")
                        shape.Text.Story.Text = new_text
                    if "BD серияКЗ" in text:
                        new_text = text.replace("BD серияКЗ", f"MD {kaz[0]}")
                        shape.Text.Story.Text = new_text
                    if "серия а" in text:
                        new_text = text.replace("серия а", f"MD {kaz[0]}")
                        shape.Text.Story.Text = new_text
                    if "рега" in text:
                        new_text = text.replace("рега", f"{eng[7]}")
                        shape.Text.Story.Text = new_text
                    if "мра" in text:
                        month_eng = months_eng[extract_date_time(eng[8])[1]]
                        new_text = text.replace("мра", f"{month_eng}")
                        shape.Text.Story.Text = new_text
                    if "дпа" in text:
                        day_eng = extract_date_time(eng[8])[0]
                        new_text = text.replace("дпа", f"{day_eng}")
                        shape.Text.Story.Text = new_text
                    if "«ОП АНГЛ»" in text:
                        op_kz = format_op_name(eng[2])
                        new_text = text.replace("«ОП АНГЛ»", f"{op_kz}")
                        shape.Text.Story.Text = new_text
                    if "степень англ" in text:
                        degree_kaz = eng[3].lower()
                        new_text = text.replace("степень англ", f"{degree_kaz}")
                        shape.Text.Story.Text = new_text
                    if "фио англ" in text:
                        fio = format_fio(eng[1])
                        shape.Text.Story.Text = shape.Text.Story.Text.replace("фио англ", fio)
                        text_width = shape.SizeWidth
                        page_width = page.SizeWidth
                        shape.PositionX = ((page_width - text_width) / 2) / 2
                    if "нпа" in text or "мпа" in text or "гпа" in text:
                        updated_text = text
                        if "нпа" in updated_text:
                            updated_text = updated_text.replace("нпа", f"№{eng[4]}")
                        if "гпа" in updated_text:
                            year_eng= extract_date_time(eng[5])[2]
                            updated_text = updated_text.replace("гпа", f"{year_eng}")
                        if "дпа" in updated_text:
                            day_eng = extract_date_time(eng[5])[0]
                            updated_text = updated_text.replace("дпа", f"{day_eng}")

                        if "мпа" in updated_text:
                            month_eng = extract_date_time(eng[5])[1]
                            month_eng = months_eng[month_eng]
                            updated_text = updated_text.replace("мпа", f"{month_eng}")

                        shape.Text.Story.Text = updated_text

                    if "дпр" in text or "№нпр" in text:
                        updated_text = text
                        
                        if "дпр" in text:
                            dpr = extract_date_time1(rus[5])
                            updated_text = updated_text.replace("дпр", f"{dpr}")
                        if "№нпр" in text:
                            updated_text = updated_text.replace("№нпр", f"№{rus[4]}")
                        shape.Text.Story.Text = updated_text
                    if "фио рус" in text:
                        fio = format_fio(rus[1])
                        shape.Text.Story.Text = shape.Text.Story.Text.replace("фио рус", fio)
                        text_width = shape.SizeWidth
                        page_width = page.SizeWidth
                        shape.PositionX = ((page_width - text_width) / 2) + 2.47
                    if "степень рус" in text:
                        degree_rus = rus[3].lower()
                        new_text = text.replace("степень рус", f"{degree_rus}")
                        shape.Text.Story.Text = new_text
                    if "«ОП РУС»" in text:
                        op_kz = format_op_name(rus[2])
                        new_text = text.replace("«ОП РУС»", f"{op_kz}")
                        shape.Text.Story.Text = new_text
                        
        base_folder = r"C:\Дипломы"
        target_folder = os.path.join(base_folder, filenamexcel)

        os.makedirs(target_folder, exist_ok=True)

        filename_cdr = f"{rus[1]}.cdr"
        output_cdr_path = os.path.join(target_folder, filename_cdr)
        document.SaveAs(output_cdr_path)

        time.sleep(2)
        document.Close()

        print(f"✅ Успешно сохранено: {output_cdr_path}")
    except Exception as e:
        print(f"❌ Ошибка при обработке: {e}")
