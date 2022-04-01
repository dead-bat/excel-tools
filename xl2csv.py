import pandas as pd
import os


def parse_to_csv(source_file, out_dir):
    filename = source_file.split("/")[-1]
    xl = pd.ExcelFile(source_file)
    ec = False
    sheetsList = xl.sheet_names
    sheetCount = len(sheetsList)
    stats = {
        "wb_data": {
            "workbook": filename,
            "output_dir": out_dir,
            "error_code": ec,
            "number_of_sheets": sheetCount,
            "sheet_names": sheetsList
        },
        "sheets": []
    }

    for sheet in sheetsList:
        of_name = filename + "_" + sheet + ".csv"
        of = os.path.join(out_dir, of_name)
        try:
            df = pd.read_excel(xl, sheet)
            df.to_csv(of, index=None)
        except Exception:
            ec = True
        finally:
            stats["sheets"].append({
                "sheet_name": sheet,
                "sheet_file": of,
                "error_code": ec if ec is not None else False
            })

    if ec is not False:
        stats["wb_data"]["error_code"] = ec

    return stats
