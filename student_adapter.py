import zipfile
import xml.etree.ElementTree as ET

def student_check(path, key):
    try:
        if key not in LABS:
            raise ValueError('Choose a valid lab rubric.')
        with open(path,'rb') as handle:
            if handle.read(5)==b'%PDF-':
                raise ValueError('This is a PDF. Export the original spreadsheet as Excel (.xlsx); renaming the file does not convert it.')
        if not zipfile.is_zipfile(path):
            raise ValueError('This is not a readable .xlsx file. Export the original spreadsheet as Excel (.xlsx) and try again.')
        with zipfile.ZipFile(path) as archive:
            entries=archive.infolist()
            if (len(entries)>2048 or sum(e.file_size for e in entries)>50*1024*1024
                or any(e.file_size>20*1024*1024 or e.flag_bits&1 for e in entries)):
                raise ValueError('This workbook is too large or encrypted. Use an unprotected .xlsx file, or ask your instructor to review it.')
            if 'xl/workbook.xml' not in archive.namelist():
                raise ValueError('This file is not an Excel workbook. Export as Excel (.xlsx).')
            for entry in entries:
                if not entry.filename.startswith('xl/worksheets/') or not entry.filename.endswith('.xml'):continue
                cells=0
                with archive.open(entry) as content:
                    for _,element in ET.iterparse(content,events=('end',)):
                        if element.tag.endswith('}c'):cells+=1
                        if cells>100000:raise ValueError('This worksheet is too large for the checker. Ask your instructor to review it.')
                        element.clear()
        score,notes=score_workbook(path,key)
        return json.dumps(dict(score=score,feedback=notes,lab=key,version=ENGINE_VERSION))
    except (ValueError,zipfile.BadZipFile,ET.ParseError) as error:
        return json.dumps(dict(error=str(error)))
    except Exception:
        return json.dumps(dict(error='This workbook could not be checked reliably. Open it in Excel, recalculate, save as .xlsx, and try again. If that doesn’t help, contact your instructor.'))
