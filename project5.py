#owen lafont

#my imports
import openpyxl
import numbers

#this is my function that handles the whole process of opening the xcel file, reading the data in the rows and displaying the information needed
def examine_population_date(xcel_file_name):
    #this opens the xcel file based off of the parameter made
    workbook_file = openpyxl.load_workbook(xcel_file_name)
    worksheet = workbook_file.active

#this is the for loop that handles the whole process of displaying the country names and population numbers
    for current_row in worksheet.rows:
        #this is the row that has the country names
        country_cell = current_row[2]
        country_name = country_cell.value

        country_type = current_row[5]
        country_type_value = country_type.value
#this is how we will only display countries like you asked for and not regions. it will only take populations from places that are listed as 'country/area'
        if country_type_value == ('Country/Area'):
            #this will gather the data from column BP(2010). This is useful because it allows us to not have to count each indiviual row to figure out what number it is. A task that would be highly inneffective in massive number xcel files.
            population_cell2010 = openpyxl.utils.cell.column_index_from_string("BP")-1
            pop2010 = current_row[population_cell2010].value

            #this will gather the data from column BZ(2020). This is useful because it allows us to not have to count each indiviual row to figure out what number it is. A task that would be highly inneffective in massive number xcel files.
            population_cell2020 = openpyxl.utils.cell.column_index_from_string("BZ")-1
            pop2020 = current_row[population_cell2020].value


#this tells the program that if it is not a number, to not collect the data and keep going. helpful to ensure we don't pickup any white data
            if not isinstance(pop2010, numbers.Number):
                continue
            if not isinstance(pop2020, numbers.Number):
                continue

#this is simply the mathematical formula used to find the difference in the years. I multiplied it by 1000 since the numbers in the excel file are by the thousands
            popdifference = (pop2020 * 1000) - (pop2010 * 1000)

#this will make it show only the populations that decreased from 2010 to 2020, as shown by the 'less than 0'
            if popdifference <0:
                print(f"{country_name} decreased in population by : {abs(int(popdifference))}")



#this is the main function. You can see how I simply pulled the other function into here and assigned the xcel file with the parameter
def main():
    examine_population_date("UN_POPULATION_Data.xlsx")











main()