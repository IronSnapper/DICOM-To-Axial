# -*- coding: utf-8 -*-
"""
Created on Fri Oct 24 15:10:10 2025

@author: 092621
"""

import os
import pandas as pd

# Parent folder containing subfolders such as 034_002, 034_003, ...
parent_path = r"N:\LA_Projects\034_WALHIP\01_034_WALHIP_MasterSet_CT\034_WALHIP_SubSet_CT_MRI\034_WALHIP_SubSet_MRI"
excel_path = r"N:\LA_Projects\034_WALHIP\01_034_WALHIP_MasterSet_CT\034_WALHIP_SubSet_CT_MRI"

# List directories only
subfolders = [d for d in os.listdir(parent_path) if os.path.isdir(os.path.join(parent_path, d))]

# Create DataFrame
df = pd.DataFrame(subfolders, columns=['Subfolder'])

# Save to Excel
excel_path = os.path.join(excel_path, "subfolder_list.xlsx")
df.to_excel(excel_path, index=False)

print("Excel file created successfully at:")
print(excel_path)
