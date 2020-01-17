import argparse
import os
import glob
import pandas as pd


def rename_files_in_folder(names_spreadsheet, folder=None):
    name_dict = load_names_dict(names_spreadsheet)

    files = glob.glob('*.jpg')

    for f in files:
        if f in name_dict:
            new_name = os.path.join(os.path.dirname(f), name_dict[f]+'.jpg')
            print('Renaming', f, 'to', new_name)
            os.rename(f, new_name)
        else:
            print('Could not find corresponding name for', f)


def load_names_dict(spreadsheet_path):
    sheet_df = pd.read_excel(spreadsheet_path)
    name_dict = {}
    for _, row in sheet_df.iterrows():
        try:
            file_name = '{:d}-{:03d}.jpg'.format(int(row['Prefix']), int(row['Photo number']))
            person_name = row['Name']
        except ValueError:
            break
        name_dict[file_name] = person_name
    return name_dict


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('spreadsheet')
    args = parser.parse_args()

    rename_files_in_folder(args.spreadsheet)
