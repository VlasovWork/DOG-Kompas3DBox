# -*- coding: utf-8 -*-
#|razvertka_kompas

import pythoncom
# импортируем DispatchEx для изоляции процессов
from win32com.client import Dispatch, DispatchEx, gencache
import re
import os
import gc

pythoncom.CoInitialize()

def clean_material_name(raw_material: str) -> str:
    """ Универсально вырезает чистую марку стали из строки КОМПАС-3D. """
    if not raw_material:
        return "Неизвестный_материал"
    text = raw_material.replace('$', '').strip()
    if ';' in text:
        text = text.split(';')[-1].strip()
    if 'ГОСТ' in text:
        text = text.split('ГОСТ')[0].strip()
    text = text.rstrip(' -_')
    return text if text else "Неизвестный_материал"

def clean_part_name(raw_name: str) -> str:
    """ Очищает наименование детали от скрытых кодов форматирования КОМПАС """
    if not raw_name:
        return "Деталь"
        
    # Оставляем только нормальные символы: буквы, цифры, дефисы и пробелы
    text = raw_name.replace('$', '').strip()
    
    text = re.sub(r'[^а-яА-ЯёЁa-zA-Z0-9\s\-]', '', text)
    
    text = re.sub(r'\s+[a-z]{4,}$', '', text).strip()
    return text


def convert_to_dxf(models_data):
    # Используем DispatchEx, чтобы запустить ОТДЕЛЬНЫЙ фоновый процесс
    kompas = DispatchEx("Kompas.Application.7")
    kompas.Visible = False

    # Подключим константы API Компас
    kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

    # Подключим описание интерфейсов API5 через DispatchEx для полной изоляции
    kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    kompas_object = kompas6_api5_module.KompasObject(
        DispatchEx("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch)
    )

    # Подключим описание интерфейсов API7 к изолированному процессу
    kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    application = kompas_api7_module.IApplication(
        kompas._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch)
    )

    # Выключаем абсолютно все всплывающие окна внутри фонового процесса
    application.HideMessage = 2

    for model_path, amount, bending in models_data:
        if not os.path.exists(model_path):
            print(f"Файл не найден: {model_path}")
            continue

        # Создаем новый документ чертежа скрыто
        Documents = application.Documents
        kompas_document = Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentDrawing, False) # False = скрыть чертеж
        kompas_document_2d = kompas_api7_module.IKompasDocument2D(kompas_document)
        iDocument2D = kompas_object.ActiveDocument2D()

        # Меняем оформление
        layout_sheets = kompas_document.LayoutSheets
        layout_sheet = layout_sheets.Item(0)
        sheet_format = layout_sheet.Format
        sheet_format.FormatMultiplicity = 1
        sheet_format.VerticalOrientation = True
        sheet_format.Format = kompas6_constants.ksFormatA4
        layout_sheet.LayoutLibraryFileName = "C:\\Program Files\\ASCON\\KOMPAS-3D v23\\Sys\\GRAPHIC.LYT"
        layout_sheet.LayoutStyleNumber = 15.0
        layout_sheet.SheetType = kompas6_constants.ksDocumentSheet
        layout_sheet.Update()

        # Получение развертки
        vlm: kompas_api7_module.IViewsAndLayersManager = kompas_document_2d.ViewsAndLayersManager
        views: kompas_api7_module.IViews = vlm.Views

        view: kompas_api7_module.IAssociationView = kompas_api7_module.IAssociationView(views.Add(2))
        view.SourceFileName = model_path
        view.Unfold = True
        view.CenterLinesVisible = False
        view.BendLinesVisible = False
        view.ProjectionName = "#Развертка"
        view.Name = ""

        iViewDesignation = kompas_api7_module.IViewDesignation(view)
        iViewDesignation.ShowUnfold = False
        rc = view.Update()

        # Открываем 3D-модель по её имени параметров, скрыто и только для чтения
        doc = application.Documents.Open(PathName=model_path, Visible=False, ReadOnly=True)
        
        if doc:
            try:
                doc_3d = kompas_api7_module.IKompasDocument3D(doc)
                top_part = doc_3d.TopPart
                
                iPart7 = kompas_api7_module.IPart7(top_part)
                iSheetMetalContainer = kompas_api7_module.ISheetMetalContainer(iPart7)
                

                # Получаем листовое тело
                iSheetMetalBodies = iSheetMetalContainer.SheetMetalBodies
                SheetMetalBody = iSheetMetalBodies.SheetMetalBody(0)
                                 
                # Получаем толщину листа (в миллиметрах)
                Part_thickness = SheetMetalBody.Thickness
                


                # Считываем свойства
                Part_name = top_part.Name
                Part_material = top_part.Material
                Part_marking = top_part.Marking
                
            finally:
                # 1 — константа kdDoNotSaveChanges. Модель закроется мгновенно без окон.
                doc.Close(1)
        else:
            kompas_document.Close(1)
            continue

        Part_material = clean_material_name(Part_material)
        Part_name = clean_part_name(Part_name)

        # Очистка имени и материала от запрещенных символов
        for char in r'/\:*?"<>|':
            Part_name = Part_name.replace(char, "_")
            Part_material = Part_material.replace(char, "_")
            Part_marking = Part_marking.replace(char, "_")

        # Формируем путь
        folder_path = os.path.dirname(model_path)
        bend_text = "Гибка" if bending else ""
        new_filename = f"{Part_marking} - {Part_name}_{Part_material}_S{Part_thickness}мм_{amount}шт_{bend_text}.dxf"
        file_path = os.path.join(folder_path, new_filename)    

        # Сохраняем в DXF
        iDocument2D.ksSaveToDXF(file_path)
        
        # Сбрасываем флаг изменений чертежа через API5, чтобы КОМПАС не считал его модифицированным
        active_doc_api5 = kompas_object.ActiveDocument2D()
        if active_doc_api5:
            active_doc_api5.ksSetDocOptions(1, 0) # 1, 0 = принудительный сброс статуса изменений

        # закрываем чертеж (1 = kdDoNotSaveChanges)
        kompas_document.Close(1)
        
        

    # Закрываем исключительно НАШ созданный фоновый КОМПАС после завершения цикла
    try:
        kompas.Quit()
    except:
        pass
        
    gc.collect()

