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

    #print("current_filepath: ", CURRENT_PATH) 
    xlsx_path = None
    for filepath in os.listdir(CURRENT_PATH):
        file_name, extension = os.path.splitext(filepath)
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
def choose_image(x, y, data_logic, ax):
    for i, (x0, y0) in enumerate(zip(x, y)):
        dl = data_logic[i]
        if dl > 4:
            path = HEADS_FOLDER + "5_smile.png"
        elif dl == 4:
            path = HEADS_FOLDER + "4_sad.png"
        elif dl == 3:
            path = HEADS_FOLDER + "3_angry.png"
        elif dl == 2:
            path = HEADS_FOLDER + "2_scary2.png"
        else:
            path = HEADS_FOLDER + "1_confused.png"

        ab = AnnotationBbox(getImage(path), (x0, y0), pad=0, frameon=False)
        ax.add_artist(ab)


@to_logger
def reformat_string_if_longer_than_n_chars(string, n):
    
    if len(string) > n:
        
        split_strings = [string[index : index + n] for index in range(0, len(string), n)]
        last_string = split_strings[-1]
        #print("------------------------")
        #print(split_strings)
        #print("Original:\n",string)
        string = "".join((s + "\n") for s in split_strings[:-1])
        string += last_string
        #print("Final String:\n",string)
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

        - Marca Temporal
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

        renames = {'Marca temporal':'Date',
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
            #print("#############################")
            print("-----------------------------")
            print("Analyzing...")
            print("-----------------------------")
            for alumn_index, alumn_name in enumerate(UNIQUE_NAMES):
                if alumn_index+1 != len(UNIQUE_NAMES):
                    print("Progress: ", alumn_index+1, "/", len(UNIQUE_NAMES), end="\r")
                else:
                    print("Progress: ", alumn_index+1, "/", len(UNIQUE_NAMES))
                alumn_bootcamp = DICT_INFO_BY_ALUMN[alumn_name][0]
                txt = "What have you learnt?\n\n"
                title = (alumn_name + " | " + alumn_bootcamp) if len((alumn_name + " | " + alumn_bootcamp)) <=54 else (alumn_name + "\n" + alumn_bootcamp)
                pdf_name = alumn_bootcamp + "_" + alumn_name + ".pdf"
                pdf_name = pdf_name.replace("/","-")
                with PdfPages(PDF_PATH + pdf_name) as alumn_pdf:
                    
                    """ First Page """
                    
                    #fig = plt.figure(figsize=(8.69,11.27), frameon=False)
                    """
                    fig = plt.figure(figsize=(11.27,11.27), frameon=False)

                    ax = fig.add_axes([0, 0, 1, 1])
                    ax.axis('off')
                    #plt.title(alumn_name + " | " + alumn_bootcamp)
                    txt += "".join(("- " + reformat_string_if_longer_than_n_chars(string, 70) + "\n") for string in DICT_INFO_BY_ALUMN[alumn_name][3])
                    fig.text(0.5, 0.9, title, ha='center', size=20, family='serif' )
                    fig.text(0.1, 0.6, txt, ha='left', horizontalalignment='left', verticalalignment='center',
                    bbox=dict(facecolor='grey', alpha=0.2), size=13)
                    plt.title(title)
                    """

                    fig = plt.figure(figsize=(11.27, 8.69), constrained_layout=True, frameon=False)
                    ax = plt.subplot(111)
                    ax.axis('off')

                    index = list(DICT_INFO_BY_ALUMN[alumn_name][2].index)
                    index_, learnt = zip(*sorted(zip(index, list(DICT_INFO_BY_ALUMN[alumn_name][3].values))))
                    colLabels = ["Date", "What have you learnt?"] # <--- 1 row, 7 columns

                    cellText = []
                    max_length = 0
                    for i, date in enumerate(index_):
                        if len(learnt[i]) > max_length:
                            max_length = len(learnt[i])
                        t = reformat_string_if_longer_than_n_chars(learnt[i], 85)
                        row = [date.strftime("%Y-%m-%d"), t]
                        cellText.append(row)
                        #cellText = [["a", "b"],["c", "d"]]

                    tab = ax.table(cellText=cellText, colLabels=colLabels, bbox=[0, 0, 1, 1], cellLoc = 'left')
                    tab.auto_set_font_size(False)
                    tab.auto_set_column_width(col=max_length) # Provide integer list of columns to adjust 
                    tab.set_fontsize(13)

                    cellDict=tab.get_celld()
                    try:
                        for i in range(len(learnt)+1):
                            cellDict[(i,0)].set_width(0.12)
                            cellDict[(i,0)].set_height(0.05)
                            cellDict[(i,1)].set_width(0.9)
                            cellDict[(i,1)].set_height(0.05)
                            

                    except Exception as er:
                        print("Error", er)
                    cellDict[(0,0)].set_height(0.05)
                    cellDict[(0,1)].set_height(0.05)

                    alumn_pdf.savefig(orientation='portrait', dpi=2000)
                    
                
                    """ Second Page """
                    
                    # Only one figure and one plot
                    # Sentiment plot 
                    fig2, (ax1) = plt.subplots(figsize=(11.27, 8.69), constrained_layout=False)
                    #fig2.subplots_adjust(hspace=0.1, bottom=0.4, top=1.5)
                    #fig2.tight_layout()
                    fig2.autofmt_xdate()

                    #plt.suptitle("Sentiment & Day Score plots", horizontalalignment='center', verticalalignment='top', fontsize = 22)
                    #plt.suptitle("Sentiment & Day Score plots", x=0.5, y=.1, horizontalalignment='center', verticalalignment='top', fontsize = 22)

                    # Day_Score
                    day_score = list(DICT_INFO_BY_ALUMN[alumn_name][2].values)
                    sentiment_data = list(DICT_INFO_BY_ALUMN[alumn_name][1].values)

                    index, day_score, sentiment_data = zip(*sorted(zip(index, day_score, sentiment_data)))

                    #ax1.legend()
                    #ax1.spines["top"].set_bounds(0, 0)
                    #ax1.set_facecolor("#fffffd")
                    #ax1.grid(True)
                    
                    ax1.scatter(index, day_score, label="Day Score")

                    #for i in range(0, len(day_score), 1):
                    #    plt.plot(index[i:i+1], day_score[i:i+1], 'r--')
                    ax1.plot(index, day_score, "--", label="Day Score")
                    ax1.set(ylabel="Happyness of day", title="Day Score")
                    ax1.scatter(index, day_score, label="Day Score")
                    ax1.fmt_xdata = mdates.DateFormatter('%d-%m-%y')
                    choose_image(x=index, y=day_score, data_logic=sentiment_data, ax=ax1)

                    ax1.set_xlim([index[0] - timedelta(3), index[0] + timedelta(25)])
                    ax1.set_ylim([0, 7])

                    alumn_pdf.savefig(orientation='landscape', dpi=2000)

                global_pdf.savefig(orientation='landscape', dpi=2000)  
                plt.close("all")
                
            print("-----------------------------")
            print("Alumns analyzed!")
            print("-----------------------------")
            #print("#############################")
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