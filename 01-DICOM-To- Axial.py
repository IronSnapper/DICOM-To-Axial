# -*- coding: utf-8 -*-
"""
Created on Fri Oct 24 15:12:46 2025

@author: 092621
"""

import pydicom
import numpy as np
import os
import copy
import pandas as pd
from datetime import datetime, timedelta

# Parent paths
parent_path = r"N:\LA_Projects\034_WALHIP\01_034_WALHIP_MasterSet_CT\034_WALHIP_SubSet_CT_MRI\034_WALHIP_SubSet_MRI"
excel_path = r"N:\LA_Projects\034_WALHIP\01_034_WALHIP_MasterSet_CT\034_WALHIP_SubSet_CT_MRI"
out_parent = r"N:\LA_Projects\034_WALHIP\01_034_WALHIP_MasterSet_CT\034_WALHIP_SubSet_CT_MRI\034_WALHIP_SubSet_MRI_Axial"
os.makedirs(out_parent, exist_ok=True)

# --- Load subfolder list from Excel ---
excel_path = os.path.join(excel_path, "subfolder_list.xlsx")
df = pd.read_excel(excel_path)
subfolders = df['Subfolder'].tolist()

# Loop through subfolders
for folder in subfolders:
    dir_path = os.path.join(parent_path, folder)

    # Skip missing/invalid folder names
    if not os.path.isdir(dir_path):
        print("Skipping invalid:", folder)
        continue

    print("Processing:", folder)

    # --------- Your EXISTING code (unchanged) ---------
    slices = []
    for f in os.listdir(dir_path):
        path = os.path.join(dir_path, f)
        try:
            ds = pydicom.dcmread(path, stop_before_pixels=True)
            if hasattr(ds, 'InstanceNumber'):
                slices.append(f)
        except:
            pass

    slices = sorted(slices, key=lambda x: int(pydicom.dcmread(os.path.join(dir_path, x), stop_before_pixels=True).InstanceNumber))

    vol = []
    for sl in slices:
        sl_path = os.path.join(dir_path, sl)
        img = pydicom.dcmread(sl_path).pixel_array
        vol.append(img)

    sl0 = pydicom.dcmread(os.path.join(dir_path, slices[0]))
    content_time = sl0.ContentTime
    image_pos = sl0[0x00200032].value
    ref_slice = copy.deepcopy(sl0)

    vol = np.asarray(vol)
    vol = np.swapaxes(vol, 0, 2)
    vol = np.swapaxes(vol, 0, 1)

    sag = np.zeros((vol.shape[0], vol.shape[-1], vol.shape[1]))
    for i in range(vol.shape[1]):
        sag[:, :, i] = vol[:, i, :]

    ax = np.zeros((vol.shape[1], vol.shape[-1], vol.shape[0]))
    for i in range(vol.shape[0]):
        ax[:, :, i] = vol[i, :, :]

    slice_thickness = float(sl0[0x00180050].value)

    # Unique output folder
    out_path = os.path.join(out_parent, folder)
    os.makedirs(out_path, exist_ok=True)

    for i in range(ax.shape[-1]):
        pix_array = ax[:, :, i]
        pix_array = np.rot90(pix_array, k=3, axes=(0, 1))
        pix_array_bytes = np.asarray(pix_array, dtype='uint16').tobytes()

        sl_to_save = copy.deepcopy(ref_slice)
        sl_to_save.InstanceNumber = i + 1
        sl_to_save.Rows = ax.shape[1]
        sl_to_save.Columns = ax.shape[0]

        c_image_pos = copy.deepcopy(image_pos)
        c_image_pos[2] = image_pos[2] + slice_thickness * i
        sl_to_save[0x00200032].value = c_image_pos
        sl_to_save[0x00201041].value = float(c_image_pos[2])
        sl_to_save[0x00200037].value = [1, 0, 0, 0, 1, 0]

        sl_to_save.PixelData = pix_array_bytes

        time_obj = datetime.strptime(content_time, '%H%M%S.%f') + timedelta(seconds=0.002 * i)
        sl_to_save.ContentTime = time_obj.strftime('%H%M%S.%f')[:-3]

        name = f'slice_{i + 1}.dcm'
        sl_to_save.save_as(os.path.join(out_path, name))

    print(f"✅ Finished: {folder}")

print("\nALL axial DICOM reconstructions complete!")
