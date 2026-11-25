# USGS StreamStats & VDOT Hydrology - Fence Query

Created by Robert Ratliff.

This MicroStation 2026 tool delineates USGS StreamStats watersheds for culvert locations, clips NHDPlus streams, and computes VDOT hydrology (Rational and other methods). It draws the basins, markers, flowpaths, raindrop paths and tables in the active DGN.

The tool comes in two versions with the same features and layout:

| Edition | Version | File |
|---|---|---|
| Python (MSPy / Tkinter) | v2.2.0 | `python/USGS_StreamStats_FenceQuery.py` |
| VBA | v1.0.1 | `vba/USGS_StreamStats_VDOT.mvba` (source in `vba/src/`) |

## Install
Copy the files to your MicroStation macros folder, for example `...\WorkSets\<WorkSet>\Standards\Macros\`.

- **Python:** key in `python load "<path>\USGS_StreamStats_FenceQuery.py"`. These packages are optional: `pyproj`, `shapely`, `openpyxl` (Excel export).
- **VBA:** key in `vba load USGS_StreamStats_VDOT`, then `vba run StreamStatsShow`. The `.bas`/`.cls`/`.frm`/`.frx` files in `vba/src/` are the project's source, exported from the `.mvba`.

## Drawing levels
Everything the tool draws goes on these levels:

- `USGS_NHD_Streams_Clipped`
- `USGS_NHD_Streams_Labels`
- `USGS_NHD_Flow_Arrows`
- `USGS_StreamStats_Basin`
- `USGS_StreamStats_Point`
- `USGS_Culvert_Inlet`
- `USGS_Longest_Flowpath`
- `USGS_Raindrop_Path`
- `USGS_StreamStats_Table`
- `USGS_Fence_Boundary`

The header's **Delete All** button deletes everything on these levels and clears the fence, after asking you to confirm. MicroStation Undo restores the deleted elements.

## Data sources
The tool pulls data from these services:

- USGS StreamStats
- USGS NLDI / NHDPlus
- USGS NSS
- USDA NRCS SSURGO (Soil Data Access), with UC Davis SoilWeb as a backup

