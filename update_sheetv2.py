import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys
sys.path.append('C:/Users/fadia/Desktop/all_vs_code_projects/VS_code_projects/testing/')
sys.setrecursionlimit(3000)  # Set the recursion limit to 3000
import retrieve_prices 
import re
from itertools import islice
import time

column_dict = {num: chr(num + 64) for num in range(1, 27)} #dict to map numbers to english letters correspondingly
global_stores_dict={}
available_to_add_store_slots=True

#THINGS TO CHANGE MANUALLY:
#  my_store_column, 


class ProductClass():
    def __init__(self,product_name,product_link,stores_list,prices_list,comparison_url) -> None:
        self.product_name=product_name
        self.product_link=product_link
        self.stores_list=stores_list
        self.prices_list=prices_list
        self.comparison_url=comparison_url
        product_location=[]
        

    def update_sheet(self,sheet,product_dict,price_matches, name_matches,item_row,my_store_location):
        starting_stores_index=list(product_dict.keys()).index("חנות 1 ") + 1
        ending_stores_index=list(product_dict.keys()).index("מיקום החנות שלי הנוכחי")
        keys_list=list(product_dict.keys())
        stores_dict={key: product_dict[key] for key in product_dict if starting_stores_index-1 <= keys_list.index(key) <= ending_stores_index-1}
        global global_stores_dict
        global_stores_dict = stores_dict

        pass
        product_row=item_row+1
        my_store_name=product_dict['החנות שלי ']
        
        if my_store_name=='':
            handle_my_store(sheet,product_row,live_price="Null")
        else:
            my_store_index=name_matches.index(my_store_name)
            live_price=price_matches[my_store_index]
            handle_my_store(sheet,product_row,live_price)
        name_matches.pop(my_store_index)
        price_matches.pop(my_store_index)

        all_product_stores=sheet.row_values(product_row) # can swap with product_dict
        for store_name, live_price in zip(name_matches,price_matches):
            if store_name in all_product_stores:
                update_store_price_if_needed(sheet,store_name,product_row,live_price)
            else:
                add_store_and_price(sheet,store_name,product_row,live_price,global_stores_dict,starting_stores_index)
        handle_my_store_location(sheet,product_row,my_store_location,product_dict)

            
def handle_my_store_location(sheet,product_row,live_store_location,product_dict):

    amount_of_columns=len(product_dict)
    previous_store_location_cell=f"{column_dict[amount_of_columns]}{product_row}"
    current_store_location_cell=f"{column_dict[amount_of_columns-1]}{product_row}"
    if "\n\n" in str(product_dict['מיקום החנות שלי הנוכחי']):
        current_locations=product_dict['מיקום החנות שלי הנוכחי'].split("\n\n")[:-1] #get current stores
    else:
        current_locations=product_dict['מיקום החנות שלי הנוכחי']
    if type(live_store_location)==str:
        live_store_location=[live_store_location]
    formatted_live_store_locations=""
    for store_location in live_store_location:
        formatted_live_store_locations+="- "+str(store_location)+"\n\n"
    if formatted_live_store_locations.split("\n\n")[:-1]==current_locations:
        return
    sheet.update(current_store_location_cell,formatted_live_store_locations)
    sheet.update(previous_store_location_cell,current_locations)


    #if sheet.acell(update_position).value!=live_price:







def update_store_price_if_needed(sheet,store_name,product_row,live_price):
    #starting_stores_column='E'
    col_index=find_column_based_on_row(sheet,store_name,product_row)
    column=column_dict[col_index+1]
    update_position=f'{column}{product_row}'
    if sheet.acell(update_position).value!=live_price:
        sheet.update(update_position,live_price)
    else:
        return
    
def add_store_and_price(sheet,store_name,product_row,live_price,stores_dict,starting_stores_index):
    global global_stores_dict
    slot_available=False
    keys_list=list(stores_dict.keys())
    for index, (key, value) in enumerate(global_stores_dict.items(),start=1):
        if (index - 1) % 2 == 0:  # Check if index is even to loop just over store names
            if value=='':
                chosen_column=starting_stores_index+index-1
                store_name_position=f'{column_dict[chosen_column]}{product_row}'
                store_price_position=f'{column_dict[chosen_column+1]}{product_row}'
                sheet.update(store_name_position,store_name)
                #update store name in dict
                stores_dict[key]=store_name
                #update store price in dict
                store_price_key=keys_list[index]
                sheet.update(store_price_position,live_price)
                stores_dict[store_price_key]=live_price
                global_stores_dict=stores_dict
                slot_available=True
                break
    global available_to_add_store_slots
    if slot_available==False and available_to_add_store_slots==True:
        add_new_store_slot(sheet,stores_dict)
        add_store_and_price(sheet,store_name,product_row,live_price,stores_dict,starting_stores_index)


def handle_my_store(sheet,product_row,live_price):
    my_store_column='D'
    update_position=f'{my_store_column}{product_row}'
    if sheet.acell(update_position).value!=live_price:
        sheet.update(update_position,live_price)
    else:
        return None




def add_new_store_slot(sheet_instance,stores_dict):
    #last_store_number=get_last_store_number(product_dict)
    amount_of_stores=len(stores_dict)//2
    if amount_of_stores<10: # we dont have 10 stores yet
        values = sheet_instance.get_all_values()
        num_rows = len(values)
        num_columns = len(values[0])
        move_last2_columns(sheet_instance,num_rows,num_columns)
        new_store_name_column=column_dict[num_columns-1]
        new_store_price_column=column_dict[num_columns]
       #sheet_instance.update()
        last_store_number=get_last_store_number(stores_dict)
        new_store_name_value=f'חנות {last_store_number+1}'
        new_store_price_value=f'מחיר {last_store_number+1}'
        sheet_instance.update(f'{new_store_name_column}1',new_store_name_value)
        sheet_instance.update(f'{new_store_price_column}1',new_store_price_value)
        global global_stores_dict
        global_stores_dict[new_store_name_value]=''
        global_stores_dict[new_store_price_value]=''
    else:
        global available_to_add_store_slots
        available_to_add_store_slots=False

    



def get_last_store_number(stores_dict):
    second_last_dict_key=list(stores_dict.keys())[len(stores_dict)-2] #get last store from stores dict
    pattern = r'\d+'  # Regular expression pattern to match one or more digits
    matches = re.findall(pattern, str(second_last_dict_key))
    if matches:
        return int(matches[0])  # Convert the first match to an integer
    return None




def move_last2_columns(sheet_instance,num_rows,num_columns):

    # Get the values in the last column
    last_column_values = sheet_instance.col_values(num_columns)
    # Get the values in the second last column
    second_last_column_values = sheet_instance.col_values(num_columns-1)

    new_last_column_index=num_columns+2
    new_second_last_column_index=num_columns+1

    #cell range vars.
    column_letter=column_dict[new_second_last_column_index]
    start_row=1
    end_row=len(second_last_column_values)
    cell_range = f'{column_letter}{start_row}:{column_letter}{end_row}'

    #move new column
    second_last_column_values=[[value] for value in second_last_column_values]
    sheet_instance.update(cell_range, second_last_column_values)


    #cell range vars.
    column_letter=column_dict[new_last_column_index]
    start_row=1
    end_row=len(last_column_values)
    cell_range = f'{column_letter}{start_row}:{column_letter}{end_row}'



    #move new column
    last_column_values=[[value] for value in last_column_values]
    sheet_instance.update(cell_range, last_column_values)


    # # Copy the values to the new columns
    # sheet_instance.update_column(column_dict[num_columns+2], last_column_values)
    # sheet_instance.update_column(column_dict[num_columns+1], second_last_column_values)

    # Delete the original last two columns
    last_column_letter = chr(ord('A') + num_columns - 1)
    second_before_last_column_letter = chr(ord('A') + num_columns - 2)

    second_before_column_range = f'{second_before_last_column_letter}{1}:{second_before_last_column_letter}{num_rows}'
    last_column_range = f'{last_column_letter}{1}:{last_column_letter}{num_rows}'
    empty_list = [[""] for _ in range(num_rows)]
    pass
    # Clear the contents of the column
    sheet_instance.update(second_before_column_range, empty_list)
    sheet_instance.update(last_column_range, empty_list)



def find_value_indices(sheet, value):
    values = sheet.get_all_values()
    for row_index, row in enumerate(values, start=1):
        for col_index, cell_value in enumerate(row, start=1):
            if cell_value == value:
                return col_index, row_index
    return None, None


def find_column_based_on_row(sheet, value, row_index): # find the value column based on give row
    values = sheet.get_all_values()
    row = values[row_index - 1]  # Adjusting row index to zero-based indexing
    for col_index, cell_value in enumerate(row, start=1):
        if cell_value == value:
            return col_index
    return None


def update_product_link(sheet_instance,product_name,product_url):
    col_index, row_index=find_value_indices(sheet_instance,product_name)
    link_position=f'{column_dict[col_index+1]}{row_index}'
    if sheet_instance.acell(link_position).value!=product_url:
        sheet_instance.update(link_position,product_url)


def main():
    # define the scope
    scope = ['https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive']
    # add credentials to the account
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        'google_sheets.json', scope)
    # authorize the clientsheet
    client = gspread.authorize(creds)
    # get the instance of the Spreadsheet
    sheet = client.open('shops_prices')
    # get the first sheet of the Spreadsheet
    sheet_instance = sheet.get_worksheet(0)
    # get all the records of the data
    dict_list = sheet_instance.get_all_records()
    headers_row=dict_list[0]

    req_obj = retrieve_prices.MakeRequest()
    HEADERS = retrieve_prices.HEADERS
    ORIGNAL_URL = "https://www.zap.co.il/search.aspx?keyword="
    
    # # Get the values of the first row
    # first_row_values = dict_list[0]
    # keys_list=list(first_row_values.keys())
    # stores_dict={key: first_row_values[key] for key in first_row_values if first_row_values-1 <= keys_list.index(key) <= first_row_values-1}
    # global global_stores_dict
    # global_stores_dict = stores_dict

    
    products_objects=[]
    num_of_rows=len(dict_list)
    #for index,_dict in enumerate(dict_list,start=1):
    print(num_of_rows)
    for i in range(num_of_rows):
        print(i)
        dict_list = sheet_instance.get_all_records()
        index=i+1
        _dict=dict_list[i] #dict_list[i]
        product_name=_dict["מוצר"]
        print(product_name)
        URL = ORIGNAL_URL + product_name
        # GET html content
        response = req_obj.get_request(URL, headers=HEADERS)
        # with open("test.txt", "w") as file1:
        #     file1.write(str(response.content.decode('utf-8')))
        # Extract model id from the html content
        model_number = req_obj.get_model_id(response.content)
        if model_number == None:
            print(product_name + "This item does not exist! Please be more specific.")
            continue
            #exit()

        comparison_url = "https://www.zap.co.il/model.aspx?modelid=" + model_number
        print("Product link: " + comparison_url +"\n\n\n")
        update_product_link(sheet_instance,product_name,comparison_url)
        respone = req_obj.get_request(comparison_url, HEADERS)
        # Get all the stores and the prices of the product.
        price_matches, name_matches = req_obj.get_companies_and_their_prices(respone.content.decode('utf-8'))
        names_ids=req_obj.get_stores_name_and_id(respone.content.decode('utf-8'))
        product_obj=ProductClass(product_name,comparison_url,name_matches, price_matches,comparison_url)
        my_store_name=_dict['החנות שלי ']
        store_id=names_ids[my_store_name]
        store_url="https://www.zap.co.il/clientcard.aspx?siteid="+store_id
        store_url_response=req_obj.get_request(store_url, HEADERS)
        my_store_location=req_obj.get_store_locations(store_url_response.content.decode('utf-8'))
        product_obj.update_sheet(sheet_instance,_dict,price_matches, name_matches,index,my_store_location)
        products_objects.append(product_obj)
        print("sleeping")
        time.sleep(30)

if __name__ == "__main__":
    main()