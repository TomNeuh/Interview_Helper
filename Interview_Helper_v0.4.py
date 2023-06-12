import streamlit as st
import pandas as pd
import numpy as np
import os
import re
import docx2txt
from nltk.tokenize import sent_tokenize
import base64
import requests
import openai
from tqdm import tqdm
import time

st.title("Expert & Research Interview Helper")
st.write("This is a tool to analyze your research or expert interviews. You can upload your interview transcripts and it will automatically generate interview insights (1) for parts of each interview seperated, (2) for each interview combined and (3) develop an initial data structure based on the Gioia (2004) structure. The tool uses the OpenAI API to generate the insights. If you have question, please feel free to reach out to: https://www.linkedin.com/in/niklas-geiss/")


open_api_key = st.text_input('Enter your open api key. This information is not recorded or stored in any way', type = "password")

focus_areas = st.text_input('Please describe the focus areas of your interview. This will be used to generate insights for these specific areas', type = "default")

#We want to accept a maximum of 5 interviews for now:
uploaded_files = st.file_uploader("Upload your interview transcripts as .docx files. You can currently upload a maximum of 5 files.", accept_multiple_files=True, type = "docx")

#Lets limit to 5 files for now:
if len(uploaded_files) > 5:
    st.write("Please upload a maximum of 5 interview transcripts")
    uploaded_files = uploaded_files[:5]

#Lets also catch the case where the user has not uploaded any files or has uploaded files of the wrong type:
if len(uploaded_files) == 0 or uploaded_files is None or uploaded_files[0].type != "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
    st.write("Please upload your interview transcripts as .docx files.")
    uploaded_files = None


#Create an empty list to store the transcripts of the docx from each interview in:
transcripts = []


clicked = st.button('Analyze Interviews')
docx2txt_list = []

#We need to define a function get_binary_file_downloader_html() to download files from streamlit:
def get_binary_file_downloader_html(bin_file, file_label='File'):
    with open(bin_file, 'rb') as f:
        data = f.read()
    bin_str = base64.b64encode(data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download {file_label}</a>'
    return href

#If clicked is True:
if clicked:
    #Check if the user has uploaded any files (is not none):
    if uploaded_files is not None and open_api_key is not None and focus_areas is not None:
        #Lets check how many interviews have been uploaded and creata variable to store the number of interviews:
        num_interviews = len(uploaded_files)

        #Lets convert the docx files to text and store them in a list:
        for file in uploaded_files:
            docx2txt_list.append(docx2txt.process(file))

        numbers_of_interviews = []

        #Now we want to loop through each interview text and extract the transcript from it, but we need to split each interview into sentences first and also need to ensure the string does not get too large to be processed in OpenAI:
        for i in range(len(docx2txt_list)):
            numbers_of_interviews.append(i+1)
            #Split the interview into sentences:
            interview = sent_tokenize(docx2txt_list[i])
            #Create an empty list to store the parts of the interview in:
            interview_parts = []
            #Create an empty string to store the sentences in:
            interview_part = ""
            #Loop through each sentence in the interview:
            for sentence in interview:
                #Add the sentence to the interview_part string:
                interview_part += sentence
                #If the interview_part string is longer than 9,000 characters:
                if len(interview_part) > 9000:
                    #Add the interview_part string to the interview_parts list:
                    interview_parts.append(interview_part)
                    #Reset the interview_part string to empty:
                    interview_part = ""
            #Add the last interview_part string to the interview_parts list:
            interview_parts.append(interview_part)
            #Add the interview_parts list to the transcripts list:
            transcripts.append(interview_parts)

        #Lets now create a list of the transcripts with the following columns: (1) Interview Number , (2) Part Number , (3) Transcript:
        transcripts_list = []
        interview_number = []
        part_number = []
        #Loop through each interview in the transcripts list:
        for i in range(len(transcripts)):
            #Loop through each part in the interview:
            for k in range(len(transcripts[i])):
                #Append the interview number to the interview_number list:
                interview_number.append(i+1)
                #Append the part number to the part_number list:
                part_number.append(k+1)
                #Append the transcript to the transcripts_list:
                transcripts_list.append(transcripts[i][k])
        
        #Lets now create a dataframe from these three lists:
        df_transcripts = pd.DataFrame(list(zip(interview_number, part_number, transcripts_list)), columns = ["Interview_Number", "Part_Number", "Transcript"])

        ###Optional: Lets output the transcripts to an excel file and let the user download it:
        ###df_transcripts.to_excel("Interview_Transcripts.xlsx", index = False)
        ###st.write("Download your interview transcripts here:")
        ###Create a link to download the excel file:
        ###st.markdown(get_binary_file_downloader_html("Interview_Transcripts.xlsx", 'Interview_Transcripts'), unsafe_allow_html=True)

        st.info("Interview transcripts have been successfully uploaded and converted to text. Please wait while the analysis is being performed. This may take a few minutes. Please do not refresh the page or close the browser window.")
        
        #Lets loop through the transcripts now and always add the respective interview and part number (e.g. Interview 1 - Part 1:) to the beginning of each transcript in the dataframe:
        for i in range(len(df_transcripts["Transcript"])):
            df_transcripts["Transcript"][i] = "Interview "+str(df_transcripts["Interview_Number"][i])+" - Part "+str(df_transcripts["Part_Number"][i])+": "+df_transcripts["Transcript"][i]

        #Lets now create a list of the transcripts for each interview:
        interview_transcripts = []
        #Loop through each interview in the transcripts list:
        for i in range(len(df_transcripts["Transcript"])):
            interview_transcripts.append(str(df_transcripts["Transcript"][i]))

        #Lets now go into the OpenAI part and generate the insights for each interview:
        try:
            openai.api_key = open_api_key
                        
            #Let us 
            def summarize_interviews(focus_areas,interview_input):
                retries = 3
                summary = None

                while retries > 0:
                    # This time, we are only summarizing the interviews
                    messages = [
                        {"role": "system", "content": "You are a GPT-3.5-turbo-model specialized on the analysis of scientific research interviews. The interviews cover: "+str(focus_areas)+ ".Due to their length, the interviews are divided into parts and marked with interview and part numbers."},
                        {"role": "user", "content": f"Quote from the research interviews. Interviews are delivered in parts to reduce length. Do not provide context to interview, focus on quoting specific sentences from the interview that represent key findings only:{interview_input}"}
                    ]

                    
                    completion2 = openai.ChatCompletion.create(
                        model="gpt-3.5-turbo",
                        messages=messages,
                        # We want to limit the summarizes to about 150 words (which is around 200 openAI tokens). If you want longer summaries, increase the max_tokens amount
                        max_tokens=200,
                        n=1,
                        stop=None,
                        temperature=0.8
                    )

                    response_text = completion2.choices[0].message.content
                    
                    # This is our quality control check. If the API has an error and doesn't generate a summary, we will retry the review 3 times. 
                    if response_text:
                        summary = response_text
                        break
                    else:
                        retries -= 1
                        time.sleep(8)
                else:
                    summary = "Summary not available."

                time.sleep(5)

                return summary

            # After chatGPT summarizes the interview, we save the summary to a list called summaries
            summaries = []

            for i in range(len(interview_transcripts)):
                interview_input = interview_transcripts[i]
                summary = summarize_interviews(focus_areas, interview_input)
                summaries.append(summary)

            # Now we add the review summaries to the original input dataframe
            df_transcripts["Summary"] = summaries

            # Save the results to a new Excel file
            output_file = "Interviews_analyzed.xlsx"
            df_transcripts.to_excel(output_file, index=False)
          
            st.divider()
            st.subheader("Analysis for each interview part completed. Please download the results below.")
            st.write("Download the results of each part here as Excel:")
            #Create a link to download the excel file:
            st.markdown(get_binary_file_downloader_html(output_file, 'Interview_Parts_Analyzed'), unsafe_allow_html=True)

            st.info("Now proceeding with merging the parts of each interview into one summary per interview.")
            st.divider()

            try:
                number_of_interviews = len(df_transcripts['Interview_Number'].unique())

                summary_list = []
                for i in range(number_of_interviews):
                    summary_list.append(df_transcripts[df_transcripts['Interview_Number'] == 'Interview '+str(i+1)]['Summary'].values)

                #We now need to combine the summaries into one string:
                summary_string = []
                for i in range(number_of_interviews):
                    summary_string.append(' '.join(summary_list[i]))

                summary_df = pd.DataFrame({'Interview': [i+1 for i in range(number_of_interviews)], 'Summary_Input': summary_string})

                #Now we need to summarize the summaries:

                def merge_summaries(focus_areas, summary_input):
                    retries = 3
                    summary = None

                    while retries > 0:
                        # This time, we are only summarizing the interviews
                        messages = [
                            {"role": "system", "content": "You are a GPT-3.5 turbo model specializing in the analysis of scientific research interviews. Due to their length, the interviews have been divided into parts and another GPT-3.5 turbo model has already quoted for each part the most important findings on the topics of: "+str(focus_areas)+""},
                            {"role": "user", "content": f"Combine the quotes collected from the different interview sections into a collection of concept quotes the whole interview. Do not categorize them yet. Some interview sections did not contain any quotes of findings and can therefore be ignored in the summary. Please provide the quotes in a list of bullet points using in total 300 words at maximum: {summary_input}"}
                        ]

                        
                        completion2 = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",
                            messages=messages,
                            # We want to limit the summarizes to about 400 token. If you want longer summaries, increase the max_tokens amount
                            max_tokens=400,
                            n=1,
                            stop=None,
                            temperature=0.8
                        )

                        response_text = completion2.choices[0].message.content
                        
                        # This is our quality control check. If the API has an error and doesn't generate a summary, we will retry the review 3 times. 
                        if response_text:
                            summary = response_text
                            break
                        else:
                            retries -= 1
                            time.sleep(8)
                    else:
                        summary = "Summary not available."
                                
                    time.sleep(0.5)

                    return summary
    
                merged_summaries = []

                #We now need to call the function for each interview and provide the function a combined string of all interview parts each:
                for summary_input in tqdm(summary_df["Summary_Input"], desc="Processing Interview Summary"):
                    summary = merge_summaries(focus_areas, summary_input)
                    merged_summaries.append(summary)
                    time.sleep(10)

                # Now we add the review summaries to the original input dataframe
                summary_df["Merged Summary"] = merged_summaries

                # Save the results to a new Excel file
                output_file_3 = "Merged_Interview_Summaries.xlsx"
                summary_df.to_excel(output_file_3, index=False)

                st.subheader("Merging of key insights from interview summaries completed. Please download the results below.")
                st.write("Download the results as Excel:")
                #Create a link to download the excel file: 
                st.markdown(get_binary_file_downloader_html(output_file_3, 'Merged_Interview_Summaries'), unsafe_allow_html=True)

                try:
                    st.info("Now structuring the insights")

                    number_of_interviews = len(df_transcripts['Interview_Number'].unique())
                    
                    #Lets create a new merged string from the merged summaries:
                    merged_summaries_string = ""

                    for i in range(len(summary_df['Merged Summary'])):
                        merged_summaries_string += summary_df['Merged Summary'][i]

                    


                    #Lets first define the inputs for the function
                    
                    #Now we need to structure the insights:
                    def structuring(summary_input):
                        retries = 3
                        summary = None

                        while retries > 0:
                            # This time, we are only summarizing the interviews
                            messages = [
                                {"role": "system", "content": "You are a GPT-3.5 turbo model specializing in the analysis of scientific research interviews and are tasked to develop a data structure. You have a total of 750 words available for this. "},
                                {"role": "user", "content": f"Using the provided quotes, apply the Gioia data structuring method (2004) to create a data structure with the provided full quotes used as 1st order concepts, summarizing them into 2nd order themes, and aggregating them into dimensions. Example output format: Dimension 1:\n-> Theme 1:\n   --> Concept 1\n   --> Concept 2\n-> Theme 2:\n   --> Concept 3\n\nDimension 2:\n-> Theme 3:\n   --> Concept 4\n   --> Concept 5\n-> Theme 4:\n   --> Concept 6\n   -->  Concept 7\nProvide data structure indicating concept to theme and dimension.: {summary_input}"}
                            ]

                            
                            completion2 = openai.ChatCompletion.create(
                                model="gpt-3.5-turbo",
                                messages=messages,
                                # We want to limit the summarizes to about 1000 token. If you want longer summaries, increase the max_tokens amount
                                max_tokens=1000,
                                n=1,
                                stop=None,
                                temperature=0.8
                            )

                            response_text = completion2.choices[0].message.content
                            
                            # This is our quality control check. If the API has an error and doesn't generate a summary, we will retry the review 3 times. 
                            if response_text:
                                summary = response_text
                                break
                            else:
                                retries -= 1
                                time.sleep(8)
                        else:
                            summary = "Summary not available."
                                    
                        time.sleep(0.5)

                        return summary
        
                    data_structure = []
                    
                    
                    #We now need to call the function once and provide the function the combined interview string as an input:
                    structure = structuring(merged_summaries_string)
                    data_structure.append(structure)
                    time.sleep(10)

                    # Now we add the review summaries to the original input dataframe
                    data_structure_df = pd.DataFrame()

                    data_structure_df["Data Structure"] = data_structure

                    # Save the results to a new Excel file
                    output_file_4 = "Initial_Data_Structure.xlsx"
                    data_structure_df.to_excel(output_file_4, index=False)

                    st.subheader("Developing of data structure completed. Please download the results below.")
                    st.write("Download the results as Excel:")
                    #Create a link to download the excel file: 
                    st.markdown(get_binary_file_downloader_html(output_file_4, 'Intial_Data_Structure'), unsafe_allow_html=True)
                
                except:
                    st.error("An error occured. This could be because OpenAI is currently overloaded. Please try again later.")

            except:
                st.error("An error occured. This could be because OpenAI is currently overloaded. Please try again later.")

        except:
            st.error("An error occured. This could be because OpenAI is currently overloaded. Please try again later.")


    else:
        st.write("Cannot start analyis, missing required inputs or wrong inputs provided.")

