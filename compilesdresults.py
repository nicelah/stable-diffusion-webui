#!/usr/bin/env python
import argparse
import os
import re
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils import FORMULAE

# Parse command-line arguments
parser = argparse.ArgumentParser()
parser.add_argument("-i", "--input", required=True, help="path to input directory")
parser.add_argument("-o", "--output", default="output", help="path to output file")
parser.add_argument("-f", "--format", default="md", choices=["md", "csv", "xlsx"], help="output format (md, csv, or xlsx)")
args = parser.parse_args()

# Regular expression pattern to extract fields
field_pattern = re.compile(r"([\w ]+): ([^,]+)(?:,|$)")

# Write header to output file
header = []
if args.format == "xlsx":
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Data"
    row_num = 1

output_file = args.output + "." + args.format
with open(output_file, "w") as f:
    header += ["Filename"]
    for filename in os.listdir(args.input):
        if filename.endswith(".txt"):
            input_file = os.path.join(args.input, filename)

            # Extract field names from input file content using regular expressions
            with open(input_file, "r") as f2:
                content = f2.read()
                for match in field_pattern.finditer(content):
                    field_name = match.group(1)
                    if field_name not in header and field_name != "Filename" and field_name != "Image":
                        header.append(field_name)

    header += ["Image"]
    if args.format == "md":
        f.write("| " + " | ".join(header) + " |\n")
        f.write("| " + " | ".join(["---"] * (len(header) )) + " |\n")
    elif args.format == "csv":
        f.write(",".join(header) + "\n")
    elif args.format == "xlsx":
        header += ["Re-use seed?"]
        for col_num, field_name in enumerate(header, 1):
            worksheet.cell(row=row_num, column=col_num).value = field_name
        row_num += 1

# Loop over input files and write rows to output file
for filename in os.listdir(args.input):
    if filename.endswith(".txt"):
        input_file = os.path.join(args.input, filename)

        # Extract fields from input file name and content using regular expressions
        filename_match = re.match(r"(\d+)-(\d+)\.txt", filename)
        fields = {"Filename": f"{filename_match.group(1)}-{filename_match.group(2)}"}
        with open(input_file, "r") as f:
            content = f.read()
            for match in field_pattern.finditer(content):
                field_name = match.group(1)
                if field_name not in header:
                    continue
                field_value = match.group(2).strip()
                fields[field_name] = field_value

        # Add image field to fields dictionary
        if args.format == "xlsx":
            fields["Re-use seed?"] = ""
            if not header:
                print("Error: no header found")
                header = ["Filename"]
            worksheet.append([fields.get(field, "") for field in  header ])
            image_path = os.path.join(args.input, f"{filename_match.group(1)}-{filename_match.group(2)}.png")
            img = Image(image_path)
            img.width = 512
            img.height = 512
            print("trying to add image to ", f"{header[-1]}{row_num}")
            worksheet.add_image(img, f"I{row_num}")
            worksheet.column_dimensions["I"].width = 512
            worksheet.row_dimensions[row_num].height = 512
            row_num += 1
            #fields["Image"] = f'=IMAGE("{image_path}", "{filename_match.group(1)}-{filename_match.group(2)}")'
        else:
            fields["Image"] = f"![[{filename_match.group(1)}-{filename_match.group(2)}.png]]"

        print("writing...  ", [fields.get(field, "") for field in  header ])
        print("header...  ",  header )
        # Write row to output file or workbook
        if args.format == "md":
            with open(output_file, "a") as f:
                f.write("| " + " | ".join([fields.get(field, "") for field in  header]) + " |\n")
        elif args.format == "csv":
            with open(output_file, "a") as f:
                f.write(",".join([fields.get(field, "") for field in header]) + "\n")
        elif args.format == "xlsx":
            workbook.save(args.output + ".xlsx")




