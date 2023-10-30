import requests
import pandas as pd
#import re
import numpy as np
from tkinter import *
import tkinter.messagebox
from tkinter.filedialog import askopenfilename,askdirectory
###################################take file path with tkinter
#root = Tk()
root = Tk()
def onClick(): 
    tkinter.messagebox.showinfo("instructions",  """1-اختار ملف اكسيل xlsx
2-لازم يكون فية شيت BasicInfo من غير فواصل bكابيتال و i كابيتال
3-لازم يكون فيه عامود AutoNumberبدون فواصل a كابيتال n كابيتال
4-اعمل sort للشيت by AutoNumber(AutoNumber مش متكرر)
5-لازم يكون فيه عامود FullAddress بدون فواصل برضة f كابيتال A كابيتال
6-لو الaddresses كتير هتستنى شوية
7-.نتايج هتطلع في شيت اكسيل اسمه newfile.
   . هتكون مترتبة بال AutoNumber.
   . خدها كوبي حطها في ملفك.
   . هتلاقي column اسمة recheck دي addresses فيها park واللي شبهها ابقى بص عليها تاني.
   .بعد ما تقرا اقفل الويندو من الأكس                             
   . لو وقفت في حاجة ابعت وتساب
   . مينا :01228951175""")
  
# Create a Button 
button = Button(root, text="Click here for instructions", command=onClick, height=5, width=20) 
  
# Set the position of button on the top of window. 
button.pack(side='top')
root.title("PyParser by Mena Ibrahmi") 
root.geometry('500x100') 
  
root.mainloop()

#filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

#to=askdirectory()
root1 = Tk()
 
filename = askopenfilename(filetypes=[('excel', 'xlsx')]) # show an "Open" dialog box and return the path to the selected file

#to=askdirectory()
print(filename)
#print(to) 
root1.destroy()
root1.mainloop()


if filename:
    pd.set_option('display.max_colwidth',100,'display.max_columns', 1000, 'display.width', 1000, 'display.max_rows',1000)
    ###################################take file path with tkinter
    df_ads=pd.read_excel(filename,sheet_name='BasicInfo')
    #####################make the columns in lower case#######to do
    df_ads.head()
    df_ads['FullAddress']=df_ads['FullAddress'].str.replace('[@#$%&*?]','',regex=True)
    df_ads['FullAddress']=df_ads['FullAddress'].str.replace('\n',' ')
    df_ads.head()
    df_ads['AutoNumber']=[i for i in range(len(df_ads))]
    df_ads.tail(5)
    inv_ads=[]
    v_ads=[]
    for i,j in zip(df_ads['AutoNumber'],df_ads['FullAddress']):
        if str.__contains__(str.lower(j),'park') or str.__contains__(str.lower(j),'zone') or str.__contains__(str.lower(j),'area') or str.__contains__(str.lower(j),'estate')or str.__contains__(str.lower(j),'centre'):
            inv_ads.append({'AutoNumber':i,'FullAddress':j})
        else:
            v_ads.append({'AutoNumber':i,'FullAddress':j})
    def Adds_cleaning(x):
        x=x.replace(',',' ')
        x=x.lower()
        x=x.split(' ')
        if 'park' in x:
            y=x.index('park')
        
        elif 'zone' in x:
            y=x.index('zone')
            
        elif 'area'in x:
            y=x.index('area')

        elif 'estate'in x:
            y=x.index('estate')
        
        elif 'centre'in x:
            y=x.index('centre')

        else:
            y=-1


        x=x[y+1:]
        return ' '.join(x)
    inv_ads_df=pd.DataFrame(inv_ads)
    inv_ads_df

    inv_ads_df['FullAddress']=inv_ads_df['FullAddress'].apply(lambda x :Adds_cleaning(x))
    v_ads_df=pd.DataFrame(v_ads)
    v_ads_df
    inv_ads_df
    dicts=[]
    for auto_num,address in zip(v_ads_df['AutoNumber'],v_ads_df['FullAddress']):
        print('->',address)
        url=f'https://maps.googleapis.com/maps/api/geocode/json?address={address}=&key=AIzaSyCLw-pUemZjGRke8neihd0Ae_urudM081s'
        requests.get(url).ok
        r=requests.get(url).json()
        df=pd.DataFrame(r['results'][0]['address_components'])
        df['types']=df['types'].apply(','.join)
        df
        ##########################################country
        cn=df[df['types'].str.contains('country')]['long_name'].values[0]
        ##########################################state
        if len(df[df['types'].str.contains('administrative_area_level_1')])>=1:
            state=df[df['types'].str.contains('administrative_area_level_1')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_2')])>=1:
            state=df[df['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
        else:
            state='-'
        #########################################city    
        if len(df[df['types'].str.contains('^locality')])>=1:
            city=df[df['types'].str.contains('^locality')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_2')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_3')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_3')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_3')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_3')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_4')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_4')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_5')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_5')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_1')])>=1:
            city=df[df['types'].str.contains('sublocality_level_1')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_2')])>=1:
            city=df[df['types'].str.contains('sublocality_level_2')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_3')])>=1:
            city=df[df['types'].str.contains('sublocality_level_3')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_4')])>=1:
            city=df[df['types'].str.contains('sublocality_level_4')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_4')])>=1:
            city=df[df['types'].str.contains('sublocality_level_5')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('postal_town')])>=1:
            city=df[df['types'].str.contains('postal_town')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('neighborhood')])>=1:
            city=df[df['types'].str.contains('neighborhood')]['long_name'].values[0]                                        
        else:
            city='-'

        """    if city=='-':
                try:
                    sub_df=pd.DataFrame(r['results'][1]['address_components'])
                    sub_df['types']=sub_df['types'].apply(','.join)
                    if len(sub_df[sub_df['types'].str.contains('^locality')])>=1:
                        city=sub_df[sub_df['types'].str.contains('^locality')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_2')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_3')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_3')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_3')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_3')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_4')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_4')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_5')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_5')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_1')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_1')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_2')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_2')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_3')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_3')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_4')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_4')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_4')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_5')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('postal_town')])>=1:
                        city=sub_df[sub_df['types'].str.contains('postal_town')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('neighborhood')])>=1:
                        city=sub_df[sub_df['types'].str.contains('neighborhood')]['long_name'].values[0]                                        
                    else:
                        city='-'
                except IndexError:
                    city='-'"""

        ####################################################################state error handling
        if state=='-':
                url=f'https://maps.googleapis.com/maps/api/geocode/json?address={cn},{city}=&key=AIzaSyCLw-pUemZjGRke8neihd0Ae_urudM081s'
                requests.get(url).ok
                r=requests.get(url).json()
                sub_df2=pd.DataFrame(r['results'][0]['address_components'])
                sub_df2['types']=sub_df2['types'].apply(','.join)
                if len(sub_df2[sub_df2['types'].str.contains('administrative_area_level_1')])>=1:
                    state=sub_df2[sub_df2['types'].str.contains('administrative_area_level_1')]['long_name'].values[0]
                elif len(sub_df2[sub_df2['types'].str.contains('administrative_area_level_2')])>=1:
                    state=sub_df2[sub_df2['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
                else:
                    state='-'
        if state.lower() =='Shanghai' or cn.lower()=="taiwan":
            city=state

        if cn=="Singapore":
            state="Singapore"
            city="Singapore"


        ######################################################################zipcode
        if len(df[df['types'].str.contains('postal_code')])>=1:
            z_c=df[df['types'].str.contains('postal_code')]['long_name'].values[0]
        #elif len(df[df['types'].str.contains('administrative_area_level_2')])>=1:
        #    state=df[df['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
        else:
            z_c='-'
        #######################################################################streetname
        if len(df[df['types'].str.contains('route')])>=1:
            st_name=df[df['types'].str.contains('route')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('street_name')])>=1:
            st_name=df[df['types'].str.contains('street_name')]['long_name'].values[0]
        else:
            st_name='-'
        
        #######################################################################number
        if len(df[df['types'].str.contains('street_number')])>=1:
            st_number=df[df['types'].str.contains('street_number')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('premise')])>=1:
            st_number=df[df['types'].str.contains('premise')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('building_number')])>=1:
            st_number=df[df['types'].str.contains('building_number')]['long_name'].values[0]
        else:
            st_number='-'
        ############################################################################################number error handling#######to do

        ##########################################################################################lat,lng    
        df1=pd.DataFrame(r['results'][0]['geometry'])
        df1
        lat=df1.loc['lat','location']
        lng=df1.loc['lng','location']
        location_type=df1.loc['lng','location_type']
        ##############################################################################################
        #address,st_name,st_number,cn,state,city,f'{lat},{lng}',location_type,z_c
        dicts.append({'AutoNumber':auto_num,
                    'FullAddress':address,
                    'BuildingNo':st_number,
                    'Street':st_name,
                    'CountryName':cn,
                    'StateProvinceName':state,
                    'CityName':city,
                    'GPS1':f'{lat},{lng}',
                    'ZipCode':z_c})
    dicts1=[]
    for auto_num,address in zip(inv_ads_df['AutoNumber'],inv_ads_df['FullAddress']):
        print('->',address)
        url=f'https://maps.googleapis.com/maps/api/geocode/json?address={address}=&key=AIzaSyCLw-pUemZjGRke8neihd0Ae_urudM081s'
        requests.get(url).ok
        r=requests.get(url).json()
        df=pd.DataFrame(r['results'][0]['address_components'])
        df['types']=df['types'].apply(','.join)
        df
        ##########################################country
        cn=df[df['types'].str.contains('country')]['long_name'].values[0]
        ##########################################state
        if len(df[df['types'].str.contains('administrative_area_level_1')])>=1:
            state=df[df['types'].str.contains('administrative_area_level_1')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_2')])>=1:
            state=df[df['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
        else:
            state='-'
        #########################################city    
        if len(df[df['types'].str.contains('^locality')])>=1:
            city=df[df['types'].str.contains('^locality')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_2')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_3')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_3')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_3')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_3')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_4')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_4')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('administrative_area_level_5')])>=1:
            city=df[df['types'].str.contains('administrative_area_level_5')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_1')])>=1:
            city=df[df['types'].str.contains('sublocality_level_1')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_2')])>=1:
            city=df[df['types'].str.contains('sublocality_level_2')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_3')])>=1:
            city=df[df['types'].str.contains('sublocality_level_3')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_4')])>=1:
            city=df[df['types'].str.contains('sublocality_level_4')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('sublocality_level_4')])>=1:
            city=df[df['types'].str.contains('sublocality_level_5')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('postal_town')])>=1:
            city=df[df['types'].str.contains('postal_town')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('neighborhood')])>=1:
            city=df[df['types'].str.contains('neighborhood')]['long_name'].values[0]                                        
        else:
            city='-'

        """    if city=='-':
                try:
                    sub_df=pd.DataFrame(r['results'][1]['address_components'])
                    sub_df['types']=sub_df['types'].apply(','.join)
                    if len(sub_df[sub_df['types'].str.contains('^locality')])>=1:
                        city=sub_df[sub_df['types'].str.contains('^locality')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_2')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_3')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_3')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_3')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_3')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_4')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_4')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('administrative_area_level_5')])>=1:
                        city=sub_df[sub_df['types'].str.contains('administrative_area_level_5')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_1')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_1')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_2')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_2')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_3')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_3')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_4')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_4')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('sublocality_level_4')])>=1:
                        city=sub_df[sub_df['types'].str.contains('sublocality_level_5')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('postal_town')])>=1:
                        city=sub_df[sub_df['types'].str.contains('postal_town')]['long_name'].values[0]
                    elif len(sub_df[sub_df['types'].str.contains('neighborhood')])>=1:
                        city=sub_df[sub_df['types'].str.contains('neighborhood')]['long_name'].values[0]                                        
                    else:
                        city='-'
                except IndexError:
                    city='-'"""

        ####################################################################state error handling
        if state=='-':
                url=f'https://maps.googleapis.com/maps/api/geocode/json?address={cn},{city}=&key=AIzaSyCLw-pUemZjGRke8neihd0Ae_urudM081s'
                requests.get(url).ok
                r=requests.get(url).json()
                sub_df2=pd.DataFrame(r['results'][0]['address_components'])
                sub_df2['types']=sub_df2['types'].apply(','.join)
                if len(sub_df2[sub_df2['types'].str.contains('administrative_area_level_1')])>=1:
                    state=sub_df2[sub_df2['types'].str.contains('administrative_area_level_1')]['long_name'].values[0]
                elif len(sub_df2[sub_df2['types'].str.contains('administrative_area_level_2')])>=1:
                    state=sub_df2[sub_df2['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
                else:
                    state='-'
        if state.lower() =='Shanghai' or cn.lower()=="taiwan":
            city=state

        if cn=="Singapore":
            state="Singapore"
            city="Singapore"


        ######################################################################zipcode
        if len(df[df['types'].str.contains('postal_code')])>=1:
            z_c=df[df['types'].str.contains('postal_code')]['long_name'].values[0]
        #elif len(df[df['types'].str.contains('administrative_area_level_2')])>=1:
        #    state=df[df['types'].str.contains('administrative_area_level_2')]['long_name'].values[0]
        else:
            z_c='-'
        #######################################################################streetname
        if len(df[df['types'].str.contains('route')])>=1:
            st_name=df[df['types'].str.contains('route')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('street_name')])>=1:
            st_name=df[df['types'].str.contains('street_name')]['long_name'].values[0]
        else:
            st_name='-'
        
        #######################################################################number
        if len(df[df['types'].str.contains('street_number')])>=1:
            st_number=df[df['types'].str.contains('street_number')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('premise')])>=1:
            st_number=df[df['types'].str.contains('premise')]['long_name'].values[0]
        elif len(df[df['types'].str.contains('building_number')])>=1:
            st_number=df[df['types'].str.contains('building_number')]['long_name'].values[0]
        else:
            st_number='-'
        ############################################################################################number error handling#######to do

        ##########################################################################################lat,lng    
        df1=pd.DataFrame(r['results'][0]['geometry'])
        df1
        lat=df1.loc['lat','location']
        lng=df1.loc['lng','location']
        location_type=df1.loc['lng','location_type']
        ##############################################################################################
        #address,st_name,st_number,cn,state,city,f'{lat},{lng}',location_type,z_c
        dicts1.append({'AutoNumber':auto_num,
                    'FullAddress':address,
                    'BuildingNo':st_number,
                    'Street':st_name,
                    'CountryName':cn,
                    'StateProvinceName':state,
                    'CityName':city,
                    'GPS1':f'{lat},{lng}',
                    'ZipCode':z_c,
                    'recheck':'yes'})
    final_v_df=pd.DataFrame(dicts)
    final_inv_df=pd.DataFrame(dicts1)
    final_df=pd.concat([final_v_df,final_inv_df],axis=0).sort_values('AutoNumber')
    final_df.tail(5)
    final_df.to_excel('newfile.xlsx',index=False)
####################save the file to exel with diffrent name#######to do
else:
    exit()