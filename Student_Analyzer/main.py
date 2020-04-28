# Imports 
import sys, os, traceback
import pandas as pd
pd.options.display.max_columns = None
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
from datetime import datetime, timedelta
import pkg_resources.py2_warn  # This module is necessary for 'pyinstaller' when it runs in a env (clean python) 

#plt.style.use('ggplot')

# Mandatory install for xlsx files
# pip3 install xlrd

"""  @@@@@@@@@  DECORATOR  @@@@@@@@@  """

def to_logger(func):
    def wrapper_do_twice(*args, **kwargs):
        try:
            func(*args, **kwargs)
            return func(*args, **kwargs)
        except Exception as error:
            LOGGER.write_log_error(error)
    return wrapper_do_twice

"""  @@@@@@@@@  LOGGER  @@@@@@@@@  """
class Logger:
    separator = "#################################"
    header = separator*2 + "\n" + separator + " START EVENTS " + separator + "\n" + separator*2 + "\n"
    short_header = "----------EVENT---------- " + "\n"
    short_header_error = "----------ERROR---------- " + "\n"

    def __init__(self, error_path=None, writer_path=None):
        pass

    def write_log_error(self, err, info=None, force_path=None):
        exc_type, exc_obj, exc_tb = sys.exc_info()  # this is to get error line number and description.
        file_name = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]  # to get File Name.
        error_string = "ERROR : Error Msg:{},File Name : {}, Line no : {}\n".format(err, file_name,
                                                                                    exc_tb.tb_lineno)
        print(error_string)

        if force_path:
            file_log = open(force_path , "a")
        else:
            file_log = open(ERROR_LOG , "a")
        file_log.write(self.short_header_error + str(datetime.now()) + "\n\n")
        if info:
            file_log.write(str(info) + "\n\n")
        file_log.write(str(err) + "\n\n")
        if type(err) != type(""):
            ex_traceback = err.__traceback__
            tb_lines = [ line.rstrip('\n') for line in
                         traceback.format_exception(err.__class__, err, ex_traceback)]
            [file_log.write(str(line)) for line in tb_lines]
            file_log.write("\n\n")

    def write_to_logger(self, to_write, starter=False, force_path=None):
        if force_path:
            file_log = open(force_path , "a")
        else:
            file_log = open(LOG_PATH , "a")
        if starter:
            file_log.write(self.header + str(datetime.now()) + "\n\n" + str(to_write) + "\n\n")
        else:
            file_log.write(self.short_header + str(datetime.now()) + "\n\n" + str(to_write) + "\n\n")
        file_log.close()

# Create pdf folder
@to_logger
def create_folder_from_path_if_not_exist(path):
    if not os.path.exists(PDF_PATH):
        os.makedirs(PDF_PATH)
        #print("Directory " , PDF_PATH ,  " Created ")
    else:    
        #print("Directory " , PDF_PATH ,  " already exists")
        pass

@to_logger
def get_last_daily_check():
    """ List all files in the same folder """
    # For jupyter files
    #for filepath in os.listdir("."):
    # For pure .py files 

    print("current_filepath: ", CURRENT_PATH) 
    xlsx_path = None
    for filepath in os.listdir(CURRENT_PATH):
        file_name, extension = os.path.splitext(filepath)
        print(filepath)
        #LOGGER.write_to_logger("filepath: " + filepath)
        if "xlsx" in extension:
            xlsx_path = filepath
            break
    if xlsx_path:
        return xlsx_path
    else:
        raise ValueError("Excel file (xlsx) with the content of alumns not found. Please add one.")

@to_logger
def getImage(path):
    return OffsetImage(plt.imread(path))

@to_logger
def choose_image(x, y, ax):
    for x0, y0 in zip(x, y):
        if y0 > 4:
            path = HEADS_FOLDER + "5_smile.png"
        elif y0 == 4:
            path = HEADS_FOLDER + "4_sad.png"
        elif y0 == 3:
            path = HEADS_FOLDER + "3_angry.png"
        elif y0 == 2:
            path = HEADS_FOLDER + "2_scary.png"
        else:
            path = HEADS_FOLDER + "1_confused.png"

        ab = AnnotationBbox(getImage(path), (x0, y0), pad=0, frameon=False)
        ax.add_artist(ab)


@to_logger
def reformat_string_if_longer_than_n_chars(string, n):

    if len(string) > n:
        split_strings = [string[index : index + n] for index in range(0, len(string), n)]
        last_string = split_strings[-1]
        string = "".join((s + "\n") for s in split_strings[:-2])
        string += last_string
    return string

@to_logger
def sentiment_transformation(x):
    value = x.lower()
    if "contento" in value:
        return 5
    elif "triste" in value:
        return 4
    elif "enfadado" in value:
        return 3
    elif "asustado" in value:
        return 2
    else:
        return 1

def main():
    """ The main process """
    try:
        create_folder_from_path_if_not_exist(path=PDF_PATH)
        EXCEL_FILEPATH = CURRENT_PATH + "\\" + get_last_daily_check()
            
        df = pd.read_excel(EXCEL_FILEPATH)

        """ Columns to focus on:

        - Fecha
        - Nombre
        - ¿Qué programa estás cursando?
        - ¿Cómo te sientes al finalizar la jornada de clase? Puedes añadir tus propias opciones, si lo deseas.
        - ¿Qué has aprendido hoy en clase?
        - ¿Qué puntuación le darías a este día?

        Renames:

        - Date
        - Name
        - Bootcamp
        - Sentiment
        - Learnt
        - Day_Score

        """

        new_columns = ["Date", "Name", "Bootcamp", "Sentiment", "Learnt", "Day_Score"]

        renames = {'Fecha':'Date',
                'Nombre':'Name',
                '¿Qué programa estás cursando?': 'Bootcamp',
                '¿Cómo te sientes al finalizar la jornada de clase? Puedes añadir tus propias opciones, si lo deseas.':'Sentiment',
                '¿Qué has aprendido hoy en clase? ':'Learnt',
                '¿Qué puntuación le darías a este día?':'Day_Score',
                    }

        # Rename columns
        df.rename(columns=renames, inplace=True) 
        df["Sentiment"] = df["Sentiment"].apply(sentiment_transformation)
        df.set_index("Date", inplace=True)

        UNIQUE_NAMES = df.Name.unique()

        # Create dict with the values of Learnt column by alumn
        for name in UNIQUE_NAMES:
            # [Bootcamp, Sentiment, Day_Score, Learnt]
            alumn_info = []
            # Bootcamp
            alumn_info.append(max(set(list(df[df.Name == name]["Bootcamp"].values)), key = list(df[df.Name == name]["Bootcamp"].values).count))
            # Sentiment
            alumn_info.append(df[df.Name == name]["Sentiment"])
            # Day_Score
            alumn_info.append(df[df.Name == name]["Day_Score"])
            # Learnt
            alumn_info.append(df[df.Name == name]["Learnt"])
            DICT_INFO_BY_ALUMN[name] = alumn_info

        """
        # Delete unnecesary columns
        columns_to_delete = [column for column in df.columns if column not in new_columns]  
        columns_to_delete.append("Learnt")
        df.drop(columns_to_delete, axis=1, inplace=True)
        """

        with PdfPages(PDF_PATH + "1_Global_Analysis.pdf") as global_pdf:
            print("#############################")
            print("-----------------------------")
            print("Analyzing...")
            print("-----------------------------")
            for alumn_name in UNIQUE_NAMES:
                alumn_bootcamp = DICT_INFO_BY_ALUMN[alumn_name][0]
                txt = "What have you learnt?\n\n"
                title = (alumn_name + " | " + alumn_bootcamp) if len((alumn_name + " | " + alumn_bootcamp)) <=54 else (alumn_name + "\n" + alumn_bootcamp)
                pdf_name = alumn_bootcamp + "_" + alumn_name + ".pdf"
                pdf_name = pdf_name.replace("/","-")
                with PdfPages(PDF_PATH + pdf_name) as alumn_pdf:
                    
                    """ First Page """
                    fig = plt.figure(figsize=(8.69,11.27), frameon=False)

                    ax = fig.add_axes([0, 0, 1, 1])
                    ax.axis('off')
                    #plt.title(alumn_name + " | " + alumn_bootcamp)
                    txt += "".join(("- " + reformat_string_if_longer_than_n_chars(string, 70) + "\n") for string in DICT_INFO_BY_ALUMN[alumn_name][3])
                    fig.text(0.5, 0.9, title, ha='center', size=20, family='serif' )
                    fig.text(0.1, 0.6, txt, ha='left', horizontalalignment='left', verticalalignment='center',
                    bbox=dict(facecolor='grey', alpha=0.2), size=13)
                    plt.title(title)
                    alumn_pdf.savefig(orientation='portrait', dpi=600)

                    """ Second Page """

                    # Sentiment plot 
                    fig2, (ax1, ax2) = plt.subplots(2, figsize=(8.69,11.27), constrained_layout=False)
                    #fig.subplots_adjust(bottom=1.4, top=1.5)
                    fig2.subplots_adjust(hspace=-.5)
                    fig2.tight_layout()
                    fig2.autofmt_xdate()

                    plt.suptitle("Sentiment & Day Score plots", x=0.5, y=.1, horizontalalignment='center', 
                    verticalalignment='top', fontsize = 22)
            
                    index = DICT_INFO_BY_ALUMN[alumn_name][1].index
                    sentiment_data = DICT_INFO_BY_ALUMN[alumn_name][1].values

                    #ax1.legend()
                    #ax1.spines["left"].set_position(("axes", .03))
                    ax1.spines["bottom"].set_bounds(0, 0)
                    ax1.set_facecolor("#fffffc")
                    #ax1.spines["bottom"].set_linewidth(6)
                    ax1.grid(True)
                    ax1.scatter(index, sentiment_data, label="Sentiment")
                    ax1.fmt_xdata = mdates.DateFormatter('%m-%d')
                    choose_image(x=index, y=sentiment_data, ax=ax1)
                    
                    
                    ax1.set_xlim([index[0] - timedelta(3), index[0] + timedelta(23)])
                    ax1.set_ylim([0, 5])

                    # Day_Score
                    index = DICT_INFO_BY_ALUMN[alumn_name][2].index
                    day_score = DICT_INFO_BY_ALUMN[alumn_name][2].values
                    
                    #ax2.legend()
                    ax2.spines["top"].set_bounds(0, 0)
                    ax2.set_facecolor("#fffffd")
                    ax2.grid(True)
                    ax2.scatter(index, day_score, label="Day Score")
                    ax2.fmt_xdata = mdates.DateFormatter('%m-%d')
                    choose_image(x=index, y=day_score, ax=ax2)

                    ax2.set_xlim([index[0] - timedelta(3), index[0] + timedelta(23)])
                    ax2.set_ylim([0, 6])

                    alumn_pdf.savefig(orientation='portrait', dpi=900)

                global_pdf.savefig(orientation='portrait', dpi=900)  
                plt.close("all")
            print("-----------------------------")
            print("Alumns analyzed!")
            print("-----------------------------")
            print("#############################")
            sys.exit()
    except Exception as error:
        LOGGER.write_log_error(err=error, info="")

if __name__ == "__main__":
    # Global variables
    CURRENT_PATH = os.path.dirname(os.path.abspath(__file__))  # This doesn't work in jupyter
    ERROR_LOG = CURRENT_PATH + "\\errors.log"
    LOG_PATH = CURRENT_PATH + "\\elog.log"
    LOGGER = Logger()
    EXCEL_FILEPATH = None
    DICT_INFO_BY_ALUMN = {}
    PDF_PATH = CURRENT_PATH + "\\pdf_by_alum\\" 
    UNIQUE_NAMES = None  # Unique names of alumns
    HEADS_FOLDER = CURRENT_PATH + "\\heads\\" 

    paths = [
        HEADS_FOLDER + "1_confused.png",
        HEADS_FOLDER + "2_scary.png",
        HEADS_FOLDER + "3_angry.png",
        HEADS_FOLDER + "4_sad.png",
        HEADS_FOLDER + "5_smile.png"
    ]

    #LOGGER.write_to_logger("Current path: " + str(CURRENT_PATH))
    main()