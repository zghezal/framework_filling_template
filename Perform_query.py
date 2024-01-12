import pandas as pd
import numpy as np
import math

# perform query

def Perform_query(data, query_to_performed):
    # data is a dataframe
    # query_to_performed is a dataframe with all the required information to build the query on data
    print('Performing query')
    data = pd.read_csv('doc.csv', decimal=',')

    query_to_performed['Performed'] = False
    query_to_performed['Query_number'] = None
    query_to_performed['Result'] = None

    # faire la liste unique des valeurs de Target issue de query_to_performed
    target_list = query_to_performed['Target'].unique()
    filtered_list = [x for x in target_list if not (isinstance(x, float) and math.isnan(x))]
    # faire la liste unique des valeurs de
    level_list = query_to_performed['Level_1'].unique()
    filtered_list_2 = [x for x in level_list if not (isinstance(x, float) and math.isnan(x))]

    # Faire les query

    #sort by filtered_list_2
    data_by = data
    #data_by = data_by['H_Level_2'].unique() 
        # loop sur les lignes de query_to_performed. construire la valeur de result de query_to_performed quand c'est possible à partir des résultats de la query
    for index, row in query_to_performed.iterrows():
            # pourcentage de progression
        print(index/len(query_to_performed))
        if row['Performed'] == False:
                if row['Target'] in data_by.columns:
                    temp=(data['Categorie_template_1'] == row['H_Level_2'])&(data['Categorie_template_2'] == row['H_Level_1'])
                    if ~temp.any():
                        temp=(data['Categorie_template_2'] == row['H_Level_1'])
                    if temp.any():
                        query_to_performed.at[index, 'Result'] = data.loc[temp, row['Target']].sum()
                        query_to_performed.at[index, 'Performed'] = True
                     
        return query_to_performed
        
              
            
        
        
        

                       
            