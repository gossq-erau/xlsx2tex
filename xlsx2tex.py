import pandas
import os
import argparse
import sys

class TexTable(object):

    col_names = str()
    col_alignment = "r"
    body = ""
    rows = []
    midrule=False

    def __init__(self,col_names=None):
        self.col_names = col_names
        for name in col_names[1:]: self.col_alignment += "l"
        return

    def to_string(self):
        #
        msg = "% This is an auto-generated table. The following packages should be included:\n"
        msg += "% \t\\usepackage{booktabs}\n"
        msg += "% \t\\usepackage{makecell}\n"
        msg += "% \t\\usepackage{multirow}\n"
        msg += "% \t\\usepackage{graphicx}\n"

        # Header
        msg += "%\n"
        msg += "\\begin{table*}[!th]\n"
        msg += "\\centering\n"
        msg += "\\caption{}\n"
        msg += "\\label{tab:}\n"
        msg += "\\resizebox{\\textwidth}{!}{%\n"

        # Columns
        msg += "\\begin{tabular}{@{}%s@{}}\\toprule\n" % self.col_alignment
        msg += "%\n"
        for name in self.col_names: msg += "\\textbf{%s}&" % name
        msg = "%s\\\\\\midrule" % msg[:-1]



        # Body
        for n,row in enumerate(self.rows):
            msg += "\n%% Row %d\n" % (n)
            for i,cell in enumerate(row):
                if "\\makecell" in cell:
                    insert_loc = cell.find("\\makecell") + len("\\makecell")
                    cell = "%s[%s]%s" % (cell[:insert_loc],self.col_alignment[i],cell[insert_loc:])
                msg += "%s&" % cell
            msg = "%s\\\\" % msg[:-1]
            if self.midrule: msg += "\\midrule"
            continue



        # Footer
        msg += "\n%\n"
        msg += "\\bottomrule\n"
        msg += "\\end{tabular}%\n"
        msg += "}\n"
        msg += "\\end{table*}\n"
        msg += "%"

        return msg

    def write_to_file(self,_file):
        with open(_file,'w') as f: 
            f.write(self.to_string())
        return



class TexTableConverter(object):

    def __init__(self,args):
        self.word_wrap_at = args.word_wrap_at
        output_folder = args.out

        excel = pandas.ExcelFile(args.xlsx)

        # Select sheets
        self.sheets_to_parse = []
        if args.sheets == None:
            self.sheets_to_parse = [i for i in range(len(excel.sheet_names))]
            df_sheets = [excel.parse(name) for name in excel.sheet_names]
        else: 
            self.sheets_to_parse = []
            df_sheets = []
            for n in args.sheets.split(","):
                i_sheet = int(n) - 1
                self.sheets_to_parse.append(i_sheet)
                df_sheets.append(excel.parse(excel.sheet_names[i_sheet]))

        # Parse tables
        tables = [self.parse_df(df) for df in df_sheets]

        # Set alignment
        if not args.alignment == None:
            for i,col_alignment in enumerate(args.alignment.split("&")):
                tables[i].col_alignment = col_alignment

        # Set midrule flag
        for i,t in enumerate(tables):
            tables[i].midrule = args.midrule

        
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # file_names = [self.friendly_file_name(name) for name in excel.sheet_names]
        # for fn in file_names:
        for i,i_sheet in enumerate(self.sheets_to_parse):
            fn = os.path.join(output_folder,self.friendly_file_name(excel.sheet_names[i_sheet]))
            tables[i].write_to_file(fn)




    def friendly_file_name(self,s):
        # @param s: Any string
        # @return A valid file name
        fn = ""
        for c in s:
            if c in " ": fn += "_"
            elif c.isalnum() or c in "-().": fn += c
        fn += ".tex"
        return fn

    def parse_df(self,df):
        # @param df: An excel sheet as a pandas dataframe
        # @return a TexTable object
        textab = TexTable(list(df.columns))
        textab.rows = [self.parse_row(row) for row in df.iterrows()]
        return textab

    def parse_row(self,row):
        # @param row: A pandas dataframe row
        # @return a formatted LaTeX tabular row
        cell_strs = [self.parse_cell(cell) for cell in row[1]]
        return cell_strs

    def parse_cell(self,cell):
        # @param cell: A pandas dataframe cell value
        # @return A formatted LateX tabular cell
        cell_str = str(cell)
        cell_str = self.word_wrap(cell_str)
        if cell_str.lower() == "nan": return ""
        elif cell_str.count("\n") == 0: return cell_str
        return "\\makecell{%s}" % cell_str.replace("\n","\\\\")

    def word_wrap(self,line):
        # @param line: line to be wrapped
        # @return A wrapped line

        # Keep user added newlines
        lines = line.split("\n")

        ###
        newline = ""

        for l in lines:
            # Wrap each word
            heads = []
            while len(l) > self.word_wrap_at:
                # Wrap the line
                wrapped_line = self.word_wrap_nearest_word(l)

                if (wrapped_line.count("\n") == 0): break

                # Append the front to heads
                wrapped_line = wrapped_line.split("\n")
                heads.append(wrapped_line[0])

                # Prepare for next iteration
                l = wrapped_line[1]
                continue
            heads.append(l)
            
            for head in heads: newline += "%s\n" % (head)
            continue

        newline = newline.rstrip()
        return newline

    def word_wrap_nearest_word(self,line):
        # Wraps a line to the nearest word.
        # @param line: sentence to be wrapped
        # @return A wrapped string

        # If the line is shorter than the word wrap amount, then it does not need to be wrapped.
        if len(line) < self.word_wrap_at:
            return line

        # Does the word wrap happen on a space?
        if line[self.word_wrap_at] == " ":
            return '%s\n%s' % (line[:self.word_wrap_at],line[self.word_wrap_at+1:])

        # Find the position at the end of each token
        end_of_token_index = []
        i = 0
        for token in line.split(" "):
            if len(token) == 0: continue
            i += len(token) + 1
            end_of_token_index.append(i)
        end_of_token_index[-1] = end_of_token_index[-1]-1
        del i

        # Find the word where the linebreak should be
        for i,x in enumerate(end_of_token_index[:-1]):
            lower = end_of_token_index[i]
            upper = end_of_token_index[i+1]
            if (upper == self.word_wrap_at):
                return '%s\n%s' % (line[:upper],line[upper:])
            if (lower <= self.word_wrap_at and self.word_wrap_at < upper):
                return '%s\n%s' % (line[:lower],line[lower:])
        
        return line



if __name__ == "__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument("-x","--xlsx", help="Excel File.")
    parser.add_argument("-o","--out", help="(Optional) Output directory.", default="out")
    parser.add_argument("-w","--word.wrap.at", help="(Optional) # of characters to wrap long lines.",
                            dest="word_wrap_at", default=50, type=int)
    parser.add_argument("-a","--alignment", help="(Optional) Table alignment by sheet. Ex. for 3 sheets \"rll&rrr&lll\"")
    parser.add_argument("-m","--midrule", help="(Optional) Add a midrule after each row.", action="store_true")
    parser.add_argument("-s","--sheets", help="(Optional) sheets to be generated, seperated by commas. ex. \"1,4,5\"",
                            type=str)
    args = parser.parse_args()

    if args.xlsx == None:
        print("Specify a xlsx file.")
        sys.exit(0)

    print(args)

    TexTableConverter(args)
    print("Complete.")
