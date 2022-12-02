#Divya 2001CE60
#Mukul 2001CE35

import streamlit as st
import subprocess #This will be used to call psat_v3.py, an all the collected values will be passed througn this.
from datetime import datetime

#Setting the webpage title.
st.set_page_config(page_title="Team: Mukul n Divya", page_icon=":hamster:", layout="wide")

#This will prevent the re-run of whole script.
if 'replace' not in st.session_state:
    st.session_state['replace'] = False

if "filter" not in st.session_state:
    st.session_state['filter'] = False

#Creating a container, it contains greetings n header.
with st.container():
    st.subheader("Hi, :wave:")
    st.subheader("This is team Mukul n Divya")

def input_collection():
    st.markdown("<h3 style='text-align: center;'>Welcome to Frontend interface for PSAT v3.0</h1>", unsafe_allow_html=True)
    
    #Taking inputs for the initial part of PSAT_V3.py
    constant_fk2d_value = st.number_input('Enter constant_fk2d_value', step=.01, format="%.2f") #"0.00"

    multiplying_factor_value = st.number_input('Enter multiplying_factor_value', step=.01, format="%.2f") #"0.00"

    Shear_velocity_value = st.number_input('Enter Shear_velocity_value', step=.01, format="%.2f") #"0.00"
    
    #Four variables will be used in project PSAT_V3.py, so declairing/initialising them.
    corr, snr, k_value, lambda_value = 0, 0, 0.0, 0.0
    
    #Logic for filtering.
    if st.button("Filter") or st.session_state.filter:
        st.session_state['filter'] = True 
        option = st.radio("Choose filtering method: ", ('C','S','A','C & S','C & A','S & A','C & S & A','All Combine')) #Giving option using st.radio.
        #Defining step size and format also.
        if (option == 'C'):
            corr = st.number_input('Enter Threshold value of C:', step=1)
        elif (option == 'S'):
            snr = st.number_input('Enter Threshold value of S:', step=1)
        elif (option == 'A'):
            lambda_value = st.number_input('Enter Lambda value of A:', step=.01, format="%.2f")
            k_value = st.number_input('Enter k value of A:', step=.01, format="%.2f")
        elif (option == 'C & S'):
            corr = st.number_input('Enter Threshold value of C:', step=1)
            snr = st.number_input('Enter Threshold value of S:', step=1)
        elif (option == 'C & A'):
            corr = st.number_input('Enter Threshold value of C:', step=1)
            lambda_value = st.number_input('Enter Lambda value of A:', step=.01, format="%.2f")
            k_value = st.number_input('Enter k value of A:', step=.01, format="%.2f")
        elif (option == 'S & A'):
            snr = st.number_input('Enter Threshold value of S:', step=1)
            lambda_value = st.number_input('Enter Lambda value of A:', step=.01, format="%.2f")
            k_value = st.number_input('Enter k value of A:', step=.01, format="%.2f")
        elif (option == 'C & S & A') or (option == 'All Combine'):
            corr = st.number_input('Enter Threshold value of C:', step=1)
            snr = st.number_input('Enter Threshold value of S:', step=1)
            lambda_value = st.number_input('Enter Lambda value of A:', step=.01, format="%.2f")
            k_value = st.number_input('Enter k value of A:', step=.01, format="%.2f")
        
        #Interface for replace.
        if st.button("Replace") or st.session_state.replace:
            st.session_state['replace'] = True
            #Giving options using st.selectbox.
            replacement_method = st.selectbox("Choose Replacement Method From Below ", ['1. Previous Point', '2. 2*last-2nd_last', '3. Overall Mean', '4. 12_Point_Strategy', '5. Mean of Previous 2 points', '6. All Sequential', '7. All Parallel'])
            replacement_method=int(replacement_method[0])
            if st.button("Compute"):
                if replacement_method>0:
                    start_time = datetime.now()
                    with st.spinner('Computing...'): #This will temprorily change the lable of Compute button till subprocess are executed.
                        #Using Subprocess library for calling psat_v3.py. 
                        #Giving all the values collected above.
                        subprocess.run(["python", "psat_v3.py", str(constant_fk2d_value), str(multiplying_factor_value), str(Shear_velocity_value), str(option), str(corr), str(snr), str(lambda_value), str(k_value), str(replacement_method)]) 
                    end_time = datetime.now()
                    #Printing running time after completetion of process.
                    st.markdown("<h3>Done! check Results_v2.</h3>", unsafe_allow_html=True)
                    st.markdown("<h2>Run time</h2>", unsafe_allow_html=True)
                    st.write(f'Start time : {start_time.strftime("%c")}')
                    st.write(f'End time : {end_time.strftime("%c")}')
                    st.write(f'Duration : {end_time - start_time}')

input_collection()
    

