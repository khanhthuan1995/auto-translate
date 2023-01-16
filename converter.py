import win32com.client
import mimetypes, os
import click
from pathlib import Path


MIME = ("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword", "application/pdf")
def validate_file(ctx, param, value):
    if os.path.isfile(value):
        if mimetypes.guess_type(value)[0] in MIME :
            return value
        else: 
            raise click.BadParameter("Format not supported")
    else: 
        raise click.BadParameter("File does not exist")

@click.command(context_settings=dict(ignore_unknown_options=True))
@click.option('--source_file_path', prompt='Path to the file', help='File to converted', callback=validate_file)
def converter(source_file_path:str):
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False
    word.Options.ConfirmConversions = False
    file_path = source_file_path.replace("\\", "/")
    # docx_file = '{0}{1}'.format(file_path, 'x')

    docx_file = os.getcwd() + "/" + Path(source_file_path).stem + ".docx"
    print(docx_file)
    wb = word.Documents.Open(source_file_path)
    # print(wb)
    try: 
        wb.SaveAs2(docx_file, FileFormat = 16)
    except AttributeError:
        print("Please set default program to Office Word")
    finally:
        wb.Close()

if __name__ == "__main__":
    converter()
    