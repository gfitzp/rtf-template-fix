# The MIT License (MIT)
# 
# Copyright © 2015 Glenn Fitzpatrick
# 
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.


import os
import re
import base64


CHANGED_INDICATOR = '**updated - old value was'

HIGHLIGHT_COLOR = 'YELLOW'


if __name__ == "__main__":

    username = os.environ.get("USERNAME")
    os.chdir(os.path.join('C:\\', 'users', username, 'Desktop', 'RTF Templates'))

    filelist = os.listdir()

    # put all templates to be modified in an "RTF Templates" folder on your desktop
    #
    # get a list of all the files in that folder
    for file in filelist:

        # only work on .rtf files that we haven't yet modified
        if file.endswith('.rtf') and not file.startswith('modified_'):

            existing_xdo = []
        
            xdocount = 1

            fields = {}

            output_docvars = False


            # open the file
            with open(file, 'r') as template:

                print("Opening {}...".format(file))
                print()

                
                # find the highest existing <?ref:xdo0000?> reference in the file in case it already has markup
                print("Searching for existing xdo references...")
                
                for line in template:
                    if re.search(r'<\?ref:xdo\d+\?>', line):
                        regexp = re.compile(r'<\?ref:xdo(?P<xdo_num>\d+)\?>')
                        result = regexp.search(line)

                        if int(result.group('xdo_num')) not in existing_xdo:
                            existing_xdo.append(int(result.group('xdo_num')))
                        
                if existing_xdo:
                    existing_xdo.sort()
                    print("Highest existing ref:xdo :", max(existing_xdo))
                    xdocount = max(existing_xdo) + xdocount
                    print("Starting count at", xdocount)
                else:
                    print("No ref:xdo tags in the file, starting count at " + str(xdocount))


                # find each instance of a BI Publisher field (<?FOO_BAR?>) in the file
                template.seek(0)
                print("Getting list of BI Publisher fields...")
                for line in template:

                    # if we come across a tag formatted like <?FOO_BAR?>, save the field name ("FOO_BAR")
                    if re.search(r'<\?[A-Z_]+\?>', line):
                        regexp = re.compile(r'<\?(?P<field_name>[A-Z_]+)\?>')
                        result = regexp.search(line)

                        field = result.group('field_name')

                        # if we haven't already saved this field
                        if field not in fields:

                            # create the formatting string from the field name
                            formatted = "<?choose:?><?when:contains({}, '{}')?><xsl:attribute xdofo:ctx=\"inline\" name=\"background-color\">{}</xsl:attribute><?{}?><?end when?><?otherwise:?><?{}?><?end otherwise?><?end choose?>".format(field, CHANGED_INDICATOR, HIGHLIGHT_COLOR, field, field)

                            # convert the formatting string to base-64 and create the document variable we'll use to reference it
                            formatted64 = "{\*\docvar {" + str('xdo{:0>4}'.format(xdocount)) + "}{" + base64.b64encode(str.encode(formatted)).decode('UTF-8') + "}}"

                            fields[field] = [xdocount, formatted64]
                            
                            xdocount = xdocount + 1


                # replace tags with base-64 formatted code
                #
                # if we found any BI Publisher fields (no need to do this if we didn't find any...)
                if fields:

                    # create the modified output file
                    with open("modified_" + file, 'w') as output:

                        # go back to the start of the input file
                        template.seek(0)
                        
                        print("Replacing tags with base-64 formatted code...")

                        # scan through each line of the input file
                        for line in template:

                            # when we find the existing docvar line, and if we haven't already written out our list of docvars
                            # write our list of base-64 markup codes
                            if re.search(r'{\\\*\\docvar {xdo\d+}', line) and xdocount != 1 and not output_docvars:
                                
                                for field in fields:
                                    output.write(fields[field][1] + '\n')

                                output_docvars = True

                            # otherwise, if we find the '\ilfomacatchlnup' tag and haven't already written out our list of docvars
                            # split the line into two parts and write our list of base-64 markup codes after the \ilfomacatchlnup tag
                            elif re.search(r'\\ilfomacatclnup\d?', line) and not output_docvars:

                                regexp = re.compile(r'(?P<beginning>.+ilfomacatclnup\d*)(?P<end>.*)')

                                result = regexp.search(line)

                                beginning = result.group('beginning')
                                end = result.group('end')

                                output.write(beginning + '\n')

                                for field in fields:
                                    output.write(fields[field][1] + '\n')

                                if end:
                                    output.write(end + '\n')

                                output_docvars = True

                                # since we're splitting the line into <part 1 \ilfomacatclnup><part 2> and inserting our docvars in the middle
                                # continue with the next line from the template
                                continue

                            # if we find a line with a BI Publisher form field, replace it with our docvar reference
                            if re.search(r'<\?[A-Z_]+\?>', line):
                                
                                regexp = re.compile(r'<\?(?P<field_name>[A-Z_]+)\?>')
                                result = regexp.search(line)

                                field = result.group('field_name')

                                line = re.sub(r"<\?" + field + r"\?>", "<?ref:xdo" + str('{:0>4}'.format(fields[field][0])) + "?>", line)

                            # write the template line to the output file
                            if line.endswith('\r\n'):
                                output.write(line[:-1])
                            else:
                                output.write(line)

                    # close the output file
                    print()
                    print("Writing 'modified_" + file + "'")
                    output.close()

            # close the input file
            template.close()

            print()
            print("====================")
            print()

    print("Done!")
    print()
