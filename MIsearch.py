import os
import sys
import openpyxl

# main function
def main():
    # parses command line for files
    if (len(sys.argv) <= 1):
        print "Please enter a filename."
        sys.exit()

    # iterates through all files
    for file_input in sys.argv[1:]:
        print "Analyzing " + file_input

        # opens input file
        try:
            wb_input = openpyxl.load_workbook(file_input, data_only=True)
        except:
            print "Invalid filename."
            continue

        # reads from input file
        sheet_input = wb_input.active

        interval1 = []
        interval2 = []
        data = []
        for row in sheet_input:
            if (row[0].value != "Experimental"):
                data.append([row[0].value, row[1].value])
                if (row[2].value != None):
                    interval1.append([row[2].value, row[3].value, row[4].value, row[5].value, row[6].value])
                    interval2.append([row[7].value, row[8].value, row[9].value, row[10].value, row[6].value])

        interval1.sort(key = lambda x: (x[1], x[2]))
        interval2.sort(key = lambda x: (x[1], x[2]))
        data.sort(key = lambda x: x[0])

        # analyzes file
        output1 = []
        output2 = []
        max_index = len(data)

        index = 0
        charge = 0
        for interval in interval1:
            if (interval[1] != charge):
                index = 0
                charge = interval[1]
            while (index < max_index):
                if (data[index][1] == interval[1] and data[index][0] >= interval[2] and data[index][0] <= interval[3]):
                    output1.append([interval[0], interval[1], interval[4]])
                    break
                if (data[index][0] > interval[3]):
                    break
                index += 1

        index = 0
        charge = 0
        for interval in interval2:
            if (interval[1] != charge):
                index = 0
                charge = interval[1]
            while (index < max_index):
                if (data[index][1] == interval[1] and data[index][0] >= interval[2] and data[index][0] <= interval[3]):
                    output2.append([interval[0], interval[1], interval[4]])
                    break
                if (data[index][0] > interval[3]):
                    break
                index += 1

        output1.sort(key = lambda x: (x[1], x[2]))
        output2.sort(key = lambda x: (x[1], x[2]))

        # compares both lists
        output = []
        for test1 in output1:
            for test2 in output2:
                if (test1[2] == test2[2] and test1[1] == test2[1]):
                    output.append([test1[1], test1[0], test2[0], test2[2]])

        # writes to output file
        wb_output = openpyxl.Workbook()
        sheet_output = wb_output.active

        row_count = 1
        sheet_output.cell(row = row_count, column = 1).value = "Output-Charge"
        sheet_output.cell(row = row_count, column = 2).value = "Output-NAT Theoretical"
        sheet_output.cell(row = row_count, column = 3).value = "Output-SIL Theoretical"
        sheet_output.cell(row = row_count, column = 4).value = "Output-ID"

        for value in output:
            row_count += 1
            sheet_output.cell(row = row_count, column = 1).value = value[0]
            sheet_output.cell(row = row_count, column = 2).value = value[1]
            sheet_output.cell(row = row_count, column = 3).value = value[2]
            sheet_output.cell(row = row_count, column = 4).value = value[3]

        # saves output file
        file_output = os.path.splitext(file_input)[0] + "_results.xlsx"

        try:
            print "Writing " + file_output
            wb_output.save(file_output)
        except:
            print "Please close output file."
            continue

if __name__ == '__main__':
    main()