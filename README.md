![DUB](https://img.shields.io/dub/l/vibe-d.svg)
[![NSF-1550890](https://img.shields.io/badge/NSF-1550890-blue.svg)](https://nsf.gov/awardsearch/showAward?AWD_ID=1550890)

# xlsx2Tilia

This python script and associated Excel template is developed to generate XML based Tilia files (https://www.tiliait.com) for ostracode paleontological data sets.  This tool was developed to facilitate the review and upload of large ostracode datasets to Neotoma (https://www.neotomadb.org).  An Excel template allows for easy editing of metadata pertaining to all records in a given dataset, enables visualization of Tilia tables and fields, and holds the database to be processed.

The use of the Tilia program is part of the Neotoma database curation procedure for new datasets.  For a database to be reviewed, each database record is converted to a separate Tilia file.  Tilia allows for the individual manual review of a database record by a Neotoma database steward and facilitates several automated data validation procedures.  If approved, the database record is uploaded and incorporated into Neotoma.  This procedure is repeated for each database record.

The xlsx2Tilia code was developed specifically to meet the research requirements of NSF 1550890 to upload the Delorme ostracode database to Neotoma (see Acknowledgements).

## How to Use This Repository

dependencies: pandas, etree, openpyxl

The python script (`xlsx2Tilia.py`) loads the Excel workbook (database and template) located in subfolder `/data` and will iteratively process all databasse records found in the `mainTable` worksheet.  Both a Tilia file (.tlx) and XML file are exported for ease of viewing (though identical) and placed in folders `/batchTILIA` and `/batchXML` respectively.  The only worksheets for which the user will need to provide input data are `mainTable` and `contacts`.  


Excel Template Contents (`inputTemplate.xlsx`):


| Worksheet | Usage |
| ---| ---|
| mainTable | database to be converted to Tilia |
| contacts  | collector, investigator, and processor contact information and affiliation|
| xlsFormat | Tilia format parameters and flags |
| wChem     | water chemistry Tilia template |
| ostracode | ostracode Tilia template and complete species list |
| site      | site level data Tilia template |
| orig_table| not used, shows complete water chemistry table available in Tilia before fields which are auto populated in Neotoma were removed |


The use of this script can be extended to other Tilia and Neotoma paleoecologic categories. The recommended procedure is as follows: 
1. create a new Tilia file using the Tilia software program, 
2. input the required data for one database record specific to your use case and save the file,
3. create a Tilia compliant XML file by modifying the xlsx2Tilia code and Excel template to match the desired Tilia XML structure found in the new Tilia generated file.  

## Specific Application and Research Purpose

The Delorme collections includes data for more than 6,719 sites (Curry et al., 2012); however, we opted to include only sites that yielded living adult specimens the Neotoma database.  Delorme also tallied species occurrence for empty valves (dead organisms), those with only juvenile valves (or specimens dead or alive), and sites with no evidence of ostracodes.  We have presented preliminary results using the new dataset exploring results of analog reconstructions of full-glacial paleoclimate in the midwestern United States (Curry and Anderson, 2017).

### References:

* Curry, B.B. and Anderson, A.C., 2017, Full-glacial temperatures based on ostracode analogs (species and assemblages) from two North American midwestern sites, The International Research Group on Ostracoda 18th International Symposium on Ostracoda (ISO-18), University of California at Santa Barbara, CA.

* Curry, B. B., Delorme, L. D., Smith, A. J., Palmer, D. F. and Stiff, B. J., 2012, The biogeography and physicochemical characteristics of aquatic habitats of freshwater ostracodes in Canada and the United States, In: Horne, D.J., Holmes, J.A., Rodriguez-Lazaro, J. & Viehberg, F. (Eds.), Ostracoda as proxies for Quaternary Climate Change. Developments in Quaternary Science, Elsevier, Amsterdam, The Netherlands, pp. 85-115.

## Acknowledgements

We would like to thank the Canadian Museum of Nature for generously granting permission for the Delorme Ostracode Dataset to be shared through the Neotoma paleoecology database.  Dr. Denis Delorme had amassed this database in the late 1960’s and 1970’s, and the data and physical collection has been archived at the Canadian Museum of Nature Research and Collections Centre in Hull, Quebec.  Individual records can be accessed through the museum collections portal (https://nature.ca/index.php?q=en/research-collections/).  

We would like to thank the National Science Foundation for their funding of this research effort under funding opportunity NSF 1550890 (https://nsf.gov/awardsearch/showAward?AWD_ID=1550890)

![footer images Neotoma NSF and EarthCube](images/footer_logos.svg)
