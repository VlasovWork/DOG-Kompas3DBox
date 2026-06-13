# -*- coding: utf-8 -*-
#|razvertka_kompas

import pythoncom
from win32com.client import Dispatch, gencache
from pathlib import Path
import os
import gc

pythoncom.CoInitialize()

def clean_material_name(raw_material: str) -> str:
    """
    Универсально вырезает чистую марку стали из строки КОМПАС-3D.
    
    Примеры работы:
    - 'Лист$d1 ГОСТ 19904-90;08Х18Н10 ГОСТ 5582-75$' -> '08Х18Н10'
    - 'Сталь 10 ГОСТ 1050-2013'                     -> 'Сталь 10'
    - 'Ст3'                                         -> 'Ст3'
    - 'Лист$d2;Сталь 20$'                           -> 'Сталь 20'
    """
    if not raw_material:
        return "Неизвестный_материал"
    
    # Шаг 1: Очищаем строку от знаков доллара $, которые КОМПАС использует для форматирования текста
    text = raw_material.replace('$', '').strip()
    
    # Шаг 2: Если есть точка с запятой (сложный сортамент), берем только то, что ПОСЛЕ нее
    if ';' in text:
        text = text.split(';')[-1].strip()
        
    if 'ГОСТ' in text:
        text = text.split('ГОСТ')[0].strip()
        
    return text if text else "Неизвестный_материал"




def convert_to_dxf(models_data):
    # API7
    kompas = Dispatch("Kompas.Application.7")

    # Делаем компас фоновым
    kompas.Visible = False

    #  Подключим константы API Компас
    kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

    #  Подключим описание интерфейсов API5
    kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    kompas_object = kompas6_api5_module.KompasObject(Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))


    #  Подключим описание интерфейсов API7
    kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    application = kompas_api7_module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))

    # Выключаем все высплывающие окна
    application.HideMessage = 2

    for model_path, amount, bending in models_data:
        
        #  Создаем новый документ
        Documents = application.Documents
        kompas_document = Documents.AddWithDefaultSettings(kompas6_constants.ksDocumentDrawing, True)

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

        # Нейминг файла развертки
        doc = application.Documents.Open(model_path)
        doc_3d = kompas_api7_module.IKompasDocument3D(doc)
        top_part = doc_3d.TopPart
        
        if doc:
            try:
                # Инициализируем интерфейс 3D-документа
                doc_3d = kompas_api7_module.IKompasDocument3D(doc)
                top_part = doc_3d.TopPart
        
                # Считываем свойства
                Part_name = top_part.Name
                Part_material = top_part.Material
                Part_marking = top_part.Marking
            finally:
                # Этот блок выполнится ВСЕГДА, даже если чтение свойств упадет с ошибкой.
                # Это гарантирует, что КОМПАС не зависнет в памяти.
                doc.Close(False)
        else:
            # Если документ не открылся, переходим к следующему файлу в цикле
            continue


        Part_material = clean_material_name(Part_material)    

        # Очистка имени и материала от запрещенных в Windows символов (/, \, *, ?, :, <, >, |, ")
        # Если в названии материала будет строка вроде "Сталь 3 ГОСТ...", двоеточие сломает сохранение!
        for char in r'/\:*?"<>|':
            Part_name = Part_name.replace(char, "_")
            Part_material = Part_material.replace(char, "_")
            Part_marking = Part_marking.replace(char, "_")

        # 3. Формируем путь
        folder_path = os.path.dirname(model_path)
        bend_text = "Гибка" if bending else ""
        new_filename = f"{Part_marking}-{Part_name}_{Part_material}_{amount}шт_{bend_text}.dxf"
        file_path = os.path.join(folder_path, new_filename)    

        iDocument2D.ksSaveToDXF(file_path)
        
    try:
        application.Quit()
    except:
        pass   
       
        # Сначала пытаемся закрыть КОМПАС стандартным методом API
    try:
        kompas.Quit()
    except:
        pass

    gc.collect()

    os.system("taskkill /f /im KOMPAS.exe")

    

