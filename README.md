# BOM Comparator
GUI application created to make comparision of specific Bill of Materials easier. You can upload raw data from different sources: PLM, SAP, scrapped from PDFs (provided by different tool) or encoded from MMC

GUI created with PySimpleGUI, data manipulation mixed usage of Pandas and custom data structures

Currently implemented modules:
- Compare BOMs
- Mass check
- Endcode MMC (Master Model Code)

  

  ## Compare BOMs

  In order to do basically anything data needs to be loaded. It can be done from file menu.

  ### Load data, subassemblies or clear loaded data
  
  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/9412f786-4b9b-421b-abdc-0091adbb058d)

  ### Part number prompts during typing:

   ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/2f38261d-acac-41d8-ad2c-664a7952ef32)

  ### Main application window

  When both parts are picked tables are populated with parts. Comparision is made and all inconsistences are marked with red (part missing) or yellow (different quantity). Comparision is made solely based on part numbers and quantities.

  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/fa0ce3bc-ff83-4b2b-ace1-1c7cac5d0710)

  ### Subassemblies
  If subassemblies are loaded they can be either quick viewed or *extracted* to Bill of Material
  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/377be6c1-8472-49c6-9758-6e90ef91add6)

  Quickview tells you what parts are inside subassembly:
  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/d4877ce0-4c3d-4c0e-a320-16602276cfc3)

  Expand options *unpacks* subassemblies and put those parts dirrectly in Bill of Material (no changes are made on raw loaded data)
  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/fcf93be1-09c8-447b-a861-6d0358d0390e)

  If user need to expand all loaded subassemblies he can use a button from top menu:
  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/10a84bf7-b0f0-472d-add6-089c6dddc4d5)

  ## Mass check
  Use this option to mass compare multiple products
  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/1e965336-b53e-4c73-9a90-3ed8855b5080)

  Output excel file is generated where each tab corresponds to each product:
  
  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/57701916-3eba-40e3-9e2f-dfea8301f5c8)

  ## Endcode MMC
  Master Model Code is alfanumerical string that represents all unique features of product
  
  Example: AAA125AF76XXX, where for example position 4-5 (*12*) would specify feature
  
  Based on logic encoded in MMC it's possible to generate Bill of Material if such logic is known.
  Application need specific logic files, list of parts with corresponding MMC, target file and if needed which tabs from logic file needs to be skipped:
  ![image](https://github.com/dczarny10/BOM_COMPARATOR/assets/105910358/2181ae14-e0cd-42c4-afd4-43928430f2a6)





