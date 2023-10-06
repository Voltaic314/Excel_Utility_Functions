'''
Author: Logan Maupin
Date: 10/06/2023

This module will take a spreadsheet column with data in it in any size chunks,
and return formatted strings of whatever excel function you wish to use.

For example, I have a column with 540 rows of numbers in column A. I wanted to
get the average values every 30 frames so finding the average of the first 30 frames,
then the average of the second set of 30 frames after that, and so on. This script
helped me do that quickly without having to manually do the math and type out each one.
'''


def get_user_input(prompt: str, should_be_an_int: bool) -> str or int:
    '''
    This function will prompt the user with whatever you need them to input, 
    if the user gave a number, then it converts the user input to an integer. 

    Parameters: 
    prompt: a string value that you are prompting your user for when you call this function
    should_be_an_int: True or False depending on whether you want your user to give you an integer.

    Returns: str or int depending on your bool. It will return whatever the user gave as input.
    '''
    
    while True:
        user_input_value = input(prompt)

        if should_be_an_int:

            user_gave_us_a_proper_num = user_input_value.isdigit()
            if user_gave_us_a_proper_num:
                return int(user_input_value)
            
            else:
                print("Please give us a proper number only.")
                continue

        else:
            return user_input_value


def get_starting_data() -> tuple[int, int, int, str, str]:
    '''
    This function prompts the user with selected user inputs to parse the data we need
    to do the rest of the script's functionality.

    Returns: step_size integer, num_of_steps integer, starting row num integer, 
    column letter str, and excel function type str
    '''
    step_size: int = get_user_input("Please provide the integer size of your chunks \n(i.e. 30 if you want this to do something for every 30 numbers in a chunk): ", True)
    total_num_of_rows: int = get_user_input("Please provide the total number of rows: ", True)
    num_of_steps: int = total_num_of_rows // step_size # round down to lowest second

    starting_value_on_spreadsheet: int = get_user_input("Please provide the row # that your data starts at: ", True)
    column_letter: str = get_user_input("Please provide the column letter that your data is in: ", False).upper()
    type_of_excel_function: str = get_user_input("Please provide the function you wish to use in excel: ", False).upper()
    tuple_of_all_the_data_we_need = (step_size, num_of_steps, starting_value_on_spreadsheet, 
                                     column_letter, type_of_excel_function)
    return tuple_of_all_the_data_we_need


def excel_function_formatting(step_size: int, function_type: str, column_letter: str, 
                              starting_row_num: int, num_of_steps: int) -> None:
    '''
    This function just formats a string for you to paste into excel and does it X number of times.
    where X is the number of seconds long your video is. 

    Parameters: 
    framerate: integer of how big your steps are between the chunks of data, like 30.
    function_type: string of your excel function type like AVERAGE for example.
    column_letter: str of your column letter like A for example.
    starting_row_num: integer of the row number that your data starts on in that column like 1.
    num_of_steps: integer of how many steps we need to do this for. 

    Returns: None
    '''
    
    ending_range_value = starting_row_num + step_size - 1
    for i in range(num_of_steps):
        excel_average_string = f'={function_type}({column_letter}{starting_row_num}:{column_letter}{ending_range_value})'
        print(excel_average_string)
        starting_row_num += step_size
        ending_range_value += step_size


def main() -> None:
    '''
    This function will ask the user for user input, then return the formatted string
    that they can then use to copy and paste into the excel column for their function values.

    Returns: None
    '''
    user_data = get_starting_data()
    framerate = user_data[0]
    video_duration_in_seconds = user_data[1]
    starting_row_num = user_data[2]
    column_letter = user_data[3]
    function_type = user_data[4]
    excel_function_formatting(step_size=framerate, function_type=function_type, column_letter=column_letter, 
                              starting_row_num=starting_row_num, num_of_steps=video_duration_in_seconds)


if __name__ == "__main__":
    main()
