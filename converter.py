import xlrd
import csv
import progress_bar as pb
import os

# Import all the file_names in the directory
directory = input("What's the directory of the folder with the xls files? "
                  "\n[please, copy and paste here and hit enter]\n")

# deletes any spaces after pasting the directory
directory = directory.split(sep=' ')[0]

# checks out every xls file in the directory
# TODO: check if this library also works for xlsx files
files_in_directory = [file_name for file_name in os.listdir(directory) if file_name.endswith('.xls')]

for file_name in files_in_directory:
    print('\n{} is being converted.'.format(file_name))
    # opens workbook
    with xlrd.open_workbook('{}/{}'.format(directory, file_name), on_demand=True) as workbook:

        # on_demand uses memory only for the current sheet
        all_sheets = workbook.sheet_names()

        # progress bar
        j = 0
        pb.print_progress(j, len(all_sheets))

        # iterates through every sheet
        for sheet_name in all_sheets:
            worksheet = workbook.sheet_by_name(sheet_name)
            first_name = file_name.split(sep='.')[0]
            with open('{}/{}_{}.csv'.format(directory, first_name, sheet_name.replace(' ', '_')), 'w') as f:
                c = csv.writer(f)
                for r in range(worksheet.nrows):
                    c.writerow(worksheet.row_values(r))
            j += 1
            pb.print_progress(j, len(all_sheets))

print('\nThe process is finished.\n***********\n')