import os
import pandas
from generate_file import GenerateFile

dataset = []

# Folder Path
path = "files"

# Change the directory
os.chdir(path)


# Read xlsx file
def count_rows_xlsx_file(file_path):
    excel_data_df = pandas.read_excel(file_path)
    return len(excel_data_df)


def run_process():
    for file in os.listdir():
        if file.endswith(".xlsx"):
            file_path = os.path.join(os.path.dirname(__file__), path, file)
            count = count_rows_xlsx_file(file_path)
            dataset.append({"file_name": file, "count": count})

    GenerateFile.csv(dataset, "report.csv")
    GenerateFile.xlsx(dataset, "report.xlsx")


def main():
    print("started process")
    run_process()
    print("end process")


def start_command():
    print("""
 ##########    ##############     #######      ##########      ##############
##                   ##         ##      ##     ##        ##          ##
##                   ##        ##        ##    ##         ##         ##
 #########           ##        ##        ##    ##       ##           ##
         ##          ##        ############    #########             ##
         ##          ##        ##        ##    ##       ##           ##
##########           ##        ##        ##    ##         ##         ##

                                            Excel File Reader (version 0.0.1)     
        """)


def end_command():
    print("""
##########    ## ##        ##    ##########
##            ##   ##      ##    ##        ##
##            ##     ##    ##    ##         ## 
#########     ##       ##  ##    ##         ##
##            ##        ## ##    ##         ##
##            ##           ##    ##        ##
##########    ##           ##    ########## 
            """)


def main():
    start_command()
    run_process()
    end_command()


if __name__ == '__main__':
    main()
