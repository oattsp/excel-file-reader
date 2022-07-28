import pandas as ps
import os


class GenerateFile:
    @staticmethod
    def csv(dataset, filename):
        data_frame = ps.DataFrame(dataset)
        output_path = os.path.join(os.path.dirname(__file__), "output", filename)
        data_frame.to_csv(output_path, sep='|', na_rep='Unkown', float_format='%.2f', index=False)

    @staticmethod
    def xlsx(dataset, filename):
        output_path = os.path.join(os.path.dirname(__file__), "output", filename)
        with ps.ExcelWriter(path=output_path, engine='xlsxwriter') as writer:
            data_frame = ps.DataFrame(dataset)
            data_frame.to_excel(writer, sheet_name='Report', index=False)

