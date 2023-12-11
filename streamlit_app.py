import streamlit as st  # pip install streamlit
import pandas as pd  # pip install pandas
import openpyxl
import plotly.express as px  # pip install plotly-express
import base64  # Standard Python Module
from io import StringIO, BytesIO  # Standard Python Module
from PIL import Image


def generate_excel_download_link(df):
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True) 
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_html_download_link(fig):
    towrite = StringIO()
    fig.write_html(towrite, include_plotlyjs="cdn")
    towrite = BytesIO(towrite.getvalue().encode())
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:text/html;charset=utf-8;base64, {b64}" download="plot.html">Download Plot</a>'
    return st.markdown(href, unsafe_allow_html=True)


st.set_page_config(page_title='Statystyki botów')
st.title('Statystyki botów 📈')
st.subheader('Dodaj plik excel')

uploaded_file = st.file_uploader('Choose a XLSM file', type='xlsm')
if uploaded_file:
    st.markdown('---')
    df = pd.read_excel(uploaded_file, engine='openpyxl')



    option = st.selectbox(
    'Jakie statystyki',
    ('Ogólne', 'Filtrowane'))
                                    
    if option == 'Ogólne':
        st.dataframe(df)
        st.markdown('---')
        mask = df['Nazwa'] == 'SUMA'
        TabelaWszystkichTransakcji = df[~mask]
        # --- PLOT PIE CHART
        pie_chart = px.pie(TabelaWszystkichTransakcji,
                title='Wszystkie transakcje przeprocesowane przez boty',
                values='Ilość razem',
                names='Nazwa')
        st.plotly_chart(pie_chart)

        st.markdown('---')
        # --- PLOT PIE CHART
        pie_chart = px.pie(TabelaWszystkichTransakcji,
                title='Transakcje zakończone sukcesem',
                values='Ilość sukcesów',
                names='Nazwa') 
        st.plotly_chart(pie_chart)

        st.markdown('---')
        # --- PLOT PIE CHART
        pie_chart = px.pie(TabelaWszystkichTransakcji,
                title='Transakcje zakończone wyjątkiem biznesowym',
                values='Ilość biznesów',
                names='Nazwa') 
        st.plotly_chart(pie_chart)

        st.markdown('---')
        # --- PLOT PIE CHART
        pie_chart = px.pie(TabelaWszystkichTransakcji,
                title='Transakcje zakończone błędem aplikacyjnym',
                values='Ilość błędów app',
                names='Nazwa')
        st.plotly_chart(pie_chart)  

        EfektywnoscRPA = df['Efektywność RPA'].unique().tolist()
        EfektywnoscRPA_selection = st.slider('Efektywność RPA',
                        value=(0,100))

        # --- FILTER DATAFRAME BASED ON SELECTION
        
        mask1 = (TabelaWszystkichTransakcji['Efektywność RPA'].between(*EfektywnoscRPA_selection)) 
        number_of_result = TabelaWszystkichTransakcji [mask1].shape[0]
        OdfiltrowanaDT = TabelaWszystkichTransakcji [mask1]

        # --- PLOT BAR CHART
        if not OdfiltrowanaDT.empty:
                bar_chart = px.bar(OdfiltrowanaDT,
                        x='Nazwa',
                        y='Efektywność RPA',
                        text='Efektywność RPA',
                        color_discrete_sequence = ['#F63366']*len(OdfiltrowanaDT),
                        template= 'plotly_white')
                st.plotly_chart(bar_chart)


    if option == 'Filtrowane':
        # --- STREAMLIT SELECTION
        NazwaBota = df.sort_values(by='Nazwa')['Nazwa'].unique().tolist()
        NazwaBota_selection = st.multiselect('Nazwa:',
                                    NazwaBota,
                                    default=None)




        # --- FILTER DATAFRAME BASED ON SELECTION
        mask = df['Nazwa'].isin(NazwaBota_selection)
        number_of_result = df[mask].shape[0]
        OdfiltrowanaDT = df[mask]


        if not OdfiltrowanaDT.empty:
                st.dataframe(OdfiltrowanaDT)
                # --- PLOT PIE CHART
                pie_chart = px.pie(OdfiltrowanaDT,
                        title='Wszystkie transakcje przeprocesowane przez wybrane boty',
                        values='Ilość razem',
                        names='Nazwa')
                st.plotly_chart(pie_chart)  

                # --- PLOT PIE CHART
                pie_chart = px.pie(OdfiltrowanaDT,
                title='Transakcje zakończone sukcesem',
                values='Ilość sukcesów',
                names='Nazwa')
                st.plotly_chart(pie_chart)  

                # --- PLOT PIE CHART
                pie_chart = px.pie(OdfiltrowanaDT,
                        title='Transakcje zakończone wyjątkiem biznesowym',
                        values='Ilość biznesów',
                        names='Nazwa')
                st.plotly_chart(pie_chart)  

                # --- PLOT PIE CHART
                pie_chart = px.pie(OdfiltrowanaDT,
                        title='Transakcje zakończone błędem aplikacyjnym',
                        values='Ilość błędów app',
                        names='Nazwa')
                st.plotly_chart(pie_chart)  






