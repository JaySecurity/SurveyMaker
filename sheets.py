import csv

infile = open('sheets.csv', 'r')
outfile = open('survey.csv','w')

reader = csv.reader(infile)
writer = csv.writer(outfile)

next(reader)
for row in reader:
    newrow = []
    if row[0].startswith('CA836'):
        newrow.extend([row[0], '','','','','','','', 'New Hours', ''])
        writer.writerow(newrow)
        writer.writerow(['Position', 'Serial #', 'Brand', 'Description', 'Prev Pressure', 'New Pressure', 'Hot / Cold', 'Prev Outer TD', 'Current Outer TD', 'Prev Inner TD', 'Current Inner TD'])
    elif row[0].startswith('P'):
        position = row[0].split('-')
        newrow.extend([position[-1], row[1], row[2], row[3], row[5], '', '',row[8],'', row[9],''])
        writer.writerow(newrow)

infile.close
outfile.close

