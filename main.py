import os
import pandas

resultSet = []

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
            resultSet.append({"file_name": file, "count": count})

    data_frame = pandas.DataFrame(resultSet)
    output_path = os.path.join(os.path.dirname(__file__), "output", "report.csv")
    data_frame.to_csv(output_path)
    data_frame.to_csv(output_path, na_rep='Unkown')  # missing value save as Unknown
    data_frame.to_csv(output_path, float_format='%.2f')  # rounded to two decimals
    data_frame.to_csv(output_path, index=False)


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
