import streamlit as st
import openpyxl

# Define the list of questions
questions = [
    "I am the life of the party.",
    "I dont talk a lot.",
    "I feel comfortable around people.",
    "I keep in the background.",
    "I start conversations.",
    "I have little to say.",
    "I talk to a lot of different people at parties.",
    "I dont like to draw attention to myself.",
    "I dont mind being the center of attention.",
    "I am quiet around strangers.",
    "I get stressed out easily.",
    "I am relaxed most of the time.",
    "I worry about things.",
    "I seldom feel blue.",
    "I am easily disturbed.",
    "I get upset easily.",
    "I change my mood a lot.",
    "I have frequent mood swings.",
    "I get irritated easily.",
    "I often feel blue.",
    "I feel little concern for others.",
    "I am interested in people.",
    "I insult people.",
    "I sympathize with others feelings.",
    "I am not interested in other peoples problems.",
    "I have a soft heart.",
    "I am not really interested in others.",
    "I take time out for others.",
    "I feel others emotions.",
    "I make people feel at ease.",
    "I am always prepared.",
    "I leave my belongings around.",
    "I pay attention to details.",
    "I make a mess of things.",
    "I get chores done right away.",
    "I often forget to put things back in their proper place.",
    "I like order.",
    "I shirk my duties.",
    "I follow a schedule.",
    "I am exacting in my work.",
    "I have a rich vocabulary.",
    "I have difficulty understanding abstract ideas.",
    "I have a vivid imagination.",
    "I am not interested in abstract ideas.",
    "I have excellent ideas.",
    "I do not have a good imagination.",
    "I am quick to understand things.",
    "I use difficult words.",
    "I spend time reflecting on things.",
    "I am full of ideas."
]
# Create an empty list to store the user's responses
responses = []

# Define the path and name of the Excel file to store the responses
responses_file = "responses.xlsx"

# Create a function to generate the Excel file
def generate_excel_file():
    # Create a new workbook object
    wb = openpyxl.Workbook()
    
    # Get the active sheet
    ws = wb.active
    
    # Write the questions in the first row
    for i, question in enumerate(questions):
        ws.cell(row=1, column=i+1).value = question
    
    # Write the responses in the second row
    for i, response in enumerate(responses):
        ws.cell(row=2, column=i+1).value = response
    
    # Save the workbook
    wb.save(responses_file)

# Create a function to reset the responses list
def reset_responses():
    responses.clear()

# Create a dictionary to map the words to numbers
word_to_number = {"Highly Disagree": 1, "Disagree": 2, "Neutral": 3, "Agree": 4, "Highly Agree": 5}

# Create the Streamlit app
st.title("Personality Prediction App")


# Display the questions in the sidebar
for i, question in enumerate(questions):
    response = st.sidebar.radio(question, ["Highly Disagree", "Disagree", "Neutral", "Agree", "Highly Agree"], key=i)
    responses.append(word_to_number[response])

# # Add a "Submit" button to generate the Excel file
# if st.sidebar.button("Submit"):
#     generate_excel_file()
#     st.success("Responses saved to Excel file")

# Add a "Submit" and "Reset" button
col1, col2 = st.sidebar.columns(2)
if col1.button("Submit"):
    generate_excel_file()
    st.success("Responses saved to Excel file")
if col2.button("Reset"):
    reset_responses()
    st.success("Responses reset")


# Load the Excel file
workbook = openpyxl.load_workbook('responses.xlsx')

# Select the worksheet
worksheet = workbook['Sheet']

# Iterate over the cells in the first row and replace their values with new values
new_values = ['EXT1','EXT2','EXT3','EXT4','EXT5','EXT6','EXT7','EXT8','EXT9','EXT10',
              'EST1','EST2','EST3','EST4','EST5','EST6','EST7','EST8','EST9','EST10',
              'AGR1','AGR2','AGR3','AGR4','AGR5','AGR6','AGR7','AGR8','AGR9','AGR10',
              'CSN1','CSN2','CSN3','CSN4','CSN5','CSN6','CSN7','CSN8','CSN9','CSN10',
              'OPN1','OPN2','OPN3','OPN4','OPN5','OPN6','OPN7','OPN8','OPN9','OPN10']

for column in worksheet.iter_cols(min_row=1, max_row=1):
    for cell in column:
        if not cell.value:
            continue
        if cell.column > len(new_values):
            break
        cell.value = new_values[cell.column - 1]
        
# Save the modified Excel file
workbook.save('personality_result.xlsx')