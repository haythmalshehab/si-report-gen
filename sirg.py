import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.graph_objects as go
import plotly.express as px
import os
import plotly.io as plt_io
# Enable this to export as iframe
# plt_io.renderers.default = 'iframe'
# 
# 
# # Here will be the production version of this script

trello_table = pd.read_csv('./INPUT/trello_board.csv')
trello_fields_dict = {'Card Name':'T#',
                          'Card Description': 'DESC',
                          'Labels': 'LABELS',
                          'List Name': 'STATUS', 
                          'T_CRE_TS':'T_CREATION_TIMESTAMP', 
                          'T_RES_TS': 'T_RESOLUTION_TIMESTAMP'}
trello_table = trello_table[trello_fields_dict]
trello_table.rename(columns=trello_fields_dict, inplace=True)
trello_table['T_CREATION_TIMESTAMP'] = pd.to_datetime(trello_table['T_CREATION_TIMESTAMP'])
trello_table['T_RESOLUTION_TIMESTAMP'] = pd.to_datetime(trello_table['T_RESOLUTION_TIMESTAMP'])

trello_table.start_timestamp = trello_table['T_CREATION_TIMESTAMP'].min().strftime('%d/%m/%Y|%H:%M:%S')
trello_table.end_timestamp = trello_table['T_CREATION_TIMESTAMP'].max().strftime('%d/%m/%Y|%H:%M:%S')
trello_table.no_days = len(trello_table['T_CREATION_TIMESTAMP'].dt.normalize().unique())

# working minus rollforward

offset = pd.offsets.CustomBusinessHour(start='08:00', end='16:00', weekmask='Sun Mon Tue Wed Thu')
trello_table['T_ACK_TIMESTAMP'] = trello_table['T_CREATION_TIMESTAMP']  + offset
offset.rollforward(trello_table['T_ACK_TIMESTAMP'])
# offset.rollforward(trello_table['T_ACK_TIMESTAMP'])
trello_table.to_csv('./OUTPUT/trello_table.csv', sep=',')