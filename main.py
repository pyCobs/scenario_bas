import openpyxl as xl
import time
import json

from load import load_pft, dico_scenario
from write import use_scenario

# chargement des 3 feuilles utiles pour ce script
start = time.time()
wb, sheet_pft, sheet_scenario, sheet_pft_bas = load_pft()
end = time.time()

print(f"Temps de chargement : {end-start} sec")

start = time.time()
scenario = dico_scenario(sheet_scenario)
use_scenario(sheet_pft, scenario, sheet_pft_bas)
end = time.time()

print(f"Temps de traitement : {end-start} sec")

start = time.time()
wb.save(filename="PF-T bas.xlsm")
end = time.time()
print(f"Temps de sauvegarde : {end-start} sec")

wb.close()
