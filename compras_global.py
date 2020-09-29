import pandas as pd
import numpy as np
import pandas.io.sql
# import matplotlib.pyplot as plt
# import seaborn as sbn
#%% Lectura de eban
df_mseg = pd.read_excel('C:/exceldb/abraham/COMPRAS/msec_2020.xlsx',sheet_name="MSEG",usecols = "J,BO,Z,AF,AL",converters = {'Material': np.str,'Pedido': np.str})
#%%
df_mseg = df_mseg.dropna(subset=['Pedido'])

#%%
df_material = pd.read_excel('C:/exceldb/abraham/COMPRAS/materiales.xlsx',sheet_name="Sheet1",usecols = "A,B,C,D",converters = {'Material': np.str})
#%%
df_sociedad = pd.read_excel('C:/exceldb/abraham/COMPRAS/Sociedades.xlsx',sheet_name="Sheet1",usecols = "A,B,C")
#%%
df_grupo = pd.read_excel('C:/exceldb/abraham/COMPRAS/grupo_articulo.xlsx',sheet_name="Sheet1",usecols = "A,B")
#%%
df_material = pd.merge(df_material,df_grupo, on='Grupo de artículos')
df_material = df_material[(df_material['Tipo material']=='ERSA') | (df_material['Tipo material']=='ROH') ] 

#%%
df_reporte = pd.merge(df_mseg,df_material, on='Material')
#%%
df_reporte = pd.merge(df_reporte,df_sociedad, on='Sociedad')
#%%
df_reporte.to_excel('C:/exceldb/abraham/COMPRAS/ReporteDetalleCompras.xlsx')
#%%
df_mat = df_reporte.pivot_table(index='Texto breve de material',columns='GRUPO SOCIEDAD',values="Importe ML",aggfunc='sum')
#%%
df_mat = df_mat[(df_mat['CMG'] > 0) & (df_mat['PROTE'] > 0) ]
#%%
df_mat.to_excel('C:/exceldb/abraham/COMPRAS/ReporteGlobalCMGProteMaterial.xlsx')
#%%
df_temp = df_reporte.pivot_table(index='Denom.gr-artículos',columns='GRUPO SOCIEDAD',values="Importe ML",aggfunc='sum')
df_temp = df_temp[(df_temp['CMG'] > 0) & (df_temp['PROTE'] > 0) ]
df_temp.to_excel('C:/exceldb/abraham/COMPRAS/ReporteGlobalCMGProteGruppoArticulo.xlsx')

#%%
df_temp = df_reporte.pivot_table(index='Texto breve de material',columns='Nombre de la empresa',values="Importe ML",aggfunc='sum')
df_temp = pd.merge(df_mat, df_temp, on='Texto breve de material')

#df_temp = df_temp[(df_temp['CMG'] > 0) & (df_temp['PROTE'] > 0) ]
df_temp.to_excel('C:/exceldb/abraham/COMPRAS/ReporteGlobalMaterialSociedad.xlsx')

#%%
df_mat = df_reporte.pivot_table(index='Texto breve de material',columns='GRUPO SOCIEDAD',values="Importe ML",aggfunc='sum')
df_mat = df_mat.fillna(0) 
df_mat = df_mat[(df_mat['CMG'] == 0) & (df_mat['PROTE'] > 0) ]
df_temp = df_reporte.pivot_table(index='Texto breve de material',columns='Nombre de la empresa',values="Importe ML",aggfunc='sum')
df_temp = pd.merge(df_mat, df_temp, on='Texto breve de material')
df_temp.to_excel('C:/exceldb/abraham/COMPRAS/ReportePROTMaterialSociedad.xlsx')

#%%
df_mat = df_reporte.pivot_table(index='Texto breve de material',columns='GRUPO SOCIEDAD',values="Importe ML",aggfunc='sum')
df_mat = df_mat.fillna(0) 
df_mat = df_mat[(df_mat['CMG'] > 0) & (df_mat['PROTE'] == 0) ]
df_temp = df_reporte.pivot_table(index='Texto breve de material',columns='Nombre de la empresa',values="Importe ML",aggfunc='sum')
df_temp = pd.merge(df_mat, df_temp, on='Texto breve de material')
df_temp.to_excel('C:/exceldb/abraham/COMPRAS/ReporteCMGMaterialSociedad.xlsx')


#df_eban = df_eban[df_eban['Indicador de borrado']!= 'X']

# df_chdr = df_chdr.rename(columns={"Valor de objeto": "Solicitud de pedido"})
# lista = df_temp['Proveedor'].to_numpy()
# df_ekko = df_ekko[~df_ekko['Proveedor'].isin(lista)]


# df_proc_aut = pd.merge(df_proc_aut,df_ekbe, on =['Documento compras','Posición'],how='inner')
# df_proc_sin_aut = pd.merge(df_proc_sin_aut,df_ekbe, on =['Documento compras','Posición'],how='left')
# df_proc_sin_aut_sin_oc = df_proc_sin_aut[pd.isna(df_proc_sin_aut['Documento compras'])]
# df_proc_sin_aut_con_oc = df_proc_sin_aut[pd.notna(df_proc_sin_aut['Documento compras'])]
# df_proc_sin_aut_con_oc_sin_ent = df_proc_sin_aut_con_oc[pd.isna(df_proc_sin_aut_con_oc['Documento material'])].copy()
# df_proc_sin_aut_con_oc_con_ent = df_proc_sin_aut_con_oc[pd.notna(df_proc_sin_aut_con_oc['Documento material'])].copy()
# df_proc_aut['time Surt'] = (df_proc_aut['Fecha de entrada'] - df_proc_aut['Creado el']).dt.days
# df_proc_aut['time OC'] = (df_proc_aut['Creado el'] - df_proc_aut['Fecha']).dt.days
# df_proc_aut['time Aut'] = (df_proc_aut['Fecha'] - df_proc_aut['Fecha de solicitud']).dt.days
# df_proc_sin_aut_con_oc_sin_ent['time OC'] = (df_proc_sin_aut_con_oc_sin_ent['Creado el'] - df_proc_sin_aut_con_oc_sin_ent['Fecha de solicitud']).dt.days
# df_proc_sin_aut_con_oc_con_ent['time OC'] = (df_proc_sin_aut_con_oc_con_ent['Creado el'] - df_proc_sin_aut_con_oc_con_ent['Fecha de solicitud']).dt.days
# df_proc_sin_aut_con_oc_con_ent['time Surt'] = (df_proc_sin_aut_con_oc_con_ent['Fecha de entrada'] - df_proc_sin_aut_con_oc_con_ent['Creado el']).dt.days
# df_proc_sin_aut_con_oc_con_ent
# df_proc_aut_b = df_proc_aut[(df_proc_aut['time OC'] >= 0)].copy()
# df_proc_aut_m = df_proc_aut[(df_proc_aut['time OC'] < 0)]
# #df_proc_aut_b.to_excel('C:/exceldb/Rodrigo/procura_natural/df_procura_aut_b.xlsx')
# #df_proc_aut_m.to_excel('C:/exceldb/Rodrigo/procura_natural/df_procura_aut_m.xlsx')
# #df_proc_sin_aut.to_excel('C:/exceldb/Rodrigo/procura_natural/df_proc_sin_aut.xlsx')
# #df_proc_sin_aut_sin_oc.to_excel('C:/exceldb/Rodrigo/procura_natural/df_proc_sin_aut_sin_oc.xlsx')
# #df_proc_sin_aut_con_oc_sin_ent.to_excel('C:/exceldb/Rodrigo/procura_natural/df_proc_sin_aut_con_oc_sin_ent.xlsx')
# #df_proc_sin_aut_con_oc_con_ent.to_excel('C:/exceldb/Rodrigo/procura_natural/df_proc_sin_aut_con_oc_con_ent.xlsx')
# df_proc_aut_b['mesn'] = pd.DatetimeIndex(df_proc_aut_b['Fecha de solicitud']).month
# df_sociedad = pd.read_excel('C:/exceldb/Rodrigo/procura_natural/sociedad_sociedadco.xlsx',sheet_name="Sheet1",usecols = "A:B")
# df_proc_aut_b = pd.merge(df_proc_aut_b,df_sociedad,on=['Sociedad'] )
# df_sociedadco = df_proc_aut_b.groupby(["sociedad co"]).count()
# df_sociedadco = df_sociedadco.reset_index()
# #%% Grafica 01
# print(df_proc_aut_b.dtypes)
# df_proc = df_proc_aut_b.copy()
# df_proc = df_proc.set_index('Fecha de solicitud')
# df_proc_sol = df_proc[['Solicitud de pedido','sociedad co']]
# df_prote = df_proc_sol.pivot_table(index='Fecha de solicitud', columns='sociedad co',values="Solicitud de pedido",aggfunc='count')
# df_prote = df_prote.resample('M').sum()

# media = df_prote.mean()

# df_prote = df_prote.reset_index()
# s_indices = df_prote['Fecha de solicitud']
# df_prote = df_prote.set_index('Fecha de solicitud')

# df_labels = s_indices.to_frame()
# df_labels['Fecha de solicitud'] = df_labels['Fecha de solicitud'].dt.strftime('%m/%Y')

# fig, ax = plt.subplots(figsize=(16, 8))
# temp_ax = df_prote.plot.bar(title='SOLICITUDES POR SOCIEDAD CO PROTE VS CMG',ax=ax)
# ax.set_xticklabels(df_labels['Fecha de solicitud'],rotation=45)
# ax.set_ylabel('NUMERO DE SOLICITUDES',fontsize=14)
# ax.set_xlabel('MESES',fontsize=14)
# media = media.reset_index()
# med_prot = media[media['sociedad co']=='PROT'].iloc[0,1]
# med_cmg = media[media['sociedad co']=='1000'].iloc[0,1]
# ax.axhline(y=med_cmg,color="b")
# ax.axhline(y=med_prot,color="r")
# Utilidades.mostrar_valor_barra( temp_ax, plt)
# plt.grid()
# fig.savefig('C:/exceldb/Rodrigo/procura_natural/img/01 solicitudes por sociedad prote vs cmg.png', transparent=False, dpi=80, bbox_inches="tight")
# #%% Grafica 02 
# df_proc_cmg = df_proc_aut_b[df_proc_aut_b['sociedad co']=='1000']
# df_proc_cmg = df_proc_cmg.set_index('Fecha de solicitud')
# df_proc_cmg = df_proc_cmg[['Solicitud de pedido','Sociedad N']]
# df_proc_cmg = df_proc_cmg.pivot_table(index='Fecha de solicitud', columns='Sociedad N',values="Solicitud de pedido",aggfunc='count')
# df_proc_cmg = df_proc_cmg.resample('M').sum()
# df_proc_cmg = df_proc_cmg.reset_index()
# df_labels = df_proc_cmg['Fecha de solicitud']
# df_labels = df_labels.to_frame()
# df_labels['Fecha de solicitud'] = df_labels['Fecha de solicitud'].dt.strftime('%m/%Y')
# df_proc_cmg = df_proc_cmg.set_index('Fecha de solicitud')

# fig, ax = plt.subplots(figsize=(16, 8))
# temp_ax = df_proc_cmg.plot.bar(title='SOLICITUDES POR SOCIEDADES DE CMG',ax=ax)
# ax.set_xticklabels(df_labels['Fecha de solicitud'],rotation=45)
# ax.set_ylabel('NUMERO DE SOLICITUDES',fontsize=14)
# ax.set_xlabel('MESES',fontsize=14)
# Utilidades.mostrar_valor_barra( temp_ax, plt)
# plt.grid()
# fig.savefig('C:/exceldb/Rodrigo/procura_natural/img/02 solicitudes por sociedades de CMG', transparent=False, dpi=80, bbox_inches="tight")
# # #%% Grafica 03
# df_proc_cmg = df_proc_aut_b[df_proc_aut_b['sociedad co']=='1000']
# df_proc_cmg = df_proc_cmg.reset_index()
# df_proc_cmg = df_proc_cmg[["Fecha de solicitud","time Aut","time OC","time Surt"]]
# df_proc_cmg = df_proc_cmg.set_index("Fecha de solicitud")
# df_proc_cmg = df_proc_cmg.resample('M').mean()
# df_proc_cmg = df_proc_cmg.reset_index()
# df_labels = df_proc_cmg['Fecha de solicitud']
# df_fechas = df_labels.to_frame()
# df_labels['Fecha de solicitud'] = df_fechas['Fecha de solicitud'].dt.strftime('%m/%Y')

# df_proc_cmg = df_proc_cmg.set_index("Fecha de solicitud")

# fig, ax = plt.subplots(figsize=(16, 8))
# temp_ax = df_proc_cmg.plot.bar(stacked = 'True',title='PROMEDIO DE TIEMPO DE PROCESO DE COMPRAS EN CMG ENERO-JUNIO',ax=ax)
# ax.set_xticklabels(df_labels['Fecha de solicitud'],rotation=45)
# ax.set_ylabel('TIEMPO DE ATENCION',fontsize=14)
# ax.set_xlabel('MESES',fontsize=14)
# # Utilidades.mostrar_valor_barra( temp_ax, plt)
# plt.grid()
# fig.savefig('C:/exceldb/Rodrigo/procura_natural/img/03 promedio de tiempo de procesos de compras enero junio', transparent=False, dpi=80, bbox_inches="tight")

# # #%% Grafica 04
# temp = df_proc_aut_b[(df_proc_aut_b['sociedad co']=='1000') & (df_proc_aut_b['mesn']==6)]
# temp = temp[["Sociedad","time Aut","time OC","time Surt"]]
# temp = temp.groupby(["Sociedad"]).mean()
# fig, ax = plt.subplots(figsize=(16, 8))
# plt.title('SOLICITUDES POR MES POR SOCIEDAD DE CMG',fontsize=14)
# labels = ax.get_xticklabels()
# plt.setp(labels, rotation=0, horizontalalignment='center',fontsize=14)
# temp.plot.bar(stacked = 'True',title='PROMEDIO DE TIEMPO DE PROCESO DE COMPRAS EN CMG DE JUNIO',ax = ax)
# ax.set_ylabel('TIEMPO',fontsize=14)
# ax.set_xlabel('SOCIEDAD',fontsize=14)
# labels = ax.get_xticklabels()
# plt.setp(labels, rotation=0, horizontalalignment='center',fontsize=14)
# plt.grid()
# fig.savefig('C:/exceldb/Rodrigo/procura_natural/img/04 promedio de tiempo de proceso de compras en cmg junio', transparent=False, dpi=80, bbox_inches="tight")

# #%% Grafica 05
# temp = df_proc_aut_b[(df_proc_aut_b['sociedad co']=='1000') & (df_proc_aut_b['mesn']==6)]
# temp = temp[["Sociedad","time OC"]]
# temp = temp.groupby(["Sociedad"]).mean()
# fig, ax = plt.subplots(figsize=(16, 8))
# ax.set_ylabel('TIEMPO',fontsize=14)
# ax.set_xlabel('MES',fontsize=14)
# labels = ax.get_xticklabels()
# plt.setp(labels, rotation=90, horizontalalignment='right',fontsize=14)
# temp_ax = temp.plot.bar(stacked = 'True',title='PROMEDIO DE TIEMPO DE ELABORACION DE OC POR SOCIEDAD EN CMG DE JUNIO',ax = ax)
# ax.set_ylabel('TIEMPO',fontsize=14)
# ax.set_xlabel('SOCIEDAD',fontsize=14)
# labels = ax.get_xticklabels()
# plt.setp(labels, rotation=0, horizontalalignment='center',fontsize=14)
# Utilidades.mostrar_valor_barra( temp_ax, plt)
# plt.grid()
# fig.savefig('C:/exceldb/Rodrigo/procura_natural/img/05 promedio de tiempo de elaboracion de cmg junio', transparent=False, dpi=80, bbox_inches="tight")

# #%% Grafica 06
# temp = df_proc_aut_b[(df_proc_aut_b['sociedad co']=='1000') & (df_proc_aut_b['mesn']==6)]
# temp = temp[["Sociedad","time OC"]]
# y = temp.pivot(columns='Sociedad',values='time OC')
# fig, ax = plt.subplots(figsize=(16, 8))
# line_props = dict(color="r", alpha=0.9)
# bbox_props = dict(color="g", alpha=0.9, linestyle="dashdot")
# flier_props = dict(marker="o", markersize=17)
# sbn.boxplot(data=y,ax=ax)
# plt.grid()
# fig.savefig('C:/exceldb/Rodrigo/procura_natural/img/06 Tiempo de elaboracion de cmg junio', transparent=False, dpi=80, bbox_inches="tight")
