# Automating Material Flows
interfacing BIM model data (Autodesk Revit) with project management files (schedules, logs) to enhance information communication between construction contractor and material suppliers or the other project stakeholders.
1. extract BIM parametric data from Revit to Excel files or database (setting up required project parameters)
2. assigning IDs (RFID tags, WBS, element ID) to each building element
3. merge BIM data with scheduling data, group data by material needed date and categories(families, types), then calculate needed material quantities for each category on distinct needed date (here consider one PO on the same start_date of a series of WBS tasks)
--PO IDs assigned at this step after "groupby".
--distinguish between insitu material and prefabricated assemblies material. insitu material dependencies: it is likely that the same type of concrete, rebars are used for slab, wall, beam construction, delivered in palleted POs; prefabricated assmeblies: used only independently with a unique RFID tag, delivered independently.
4. generating Line of Balance charts to show the material flows(tranportation and consumption) at each day
5. write data back to BIM model, e.g., using Dynamo to populate data back to Revit.
6. update schedules and element parameters to activate the alert actions (e.g., delays of look ahead construtcion--> postpone the orgianlly planned delivery of material)
--determine the baseline schedule of delivery, baseline of material release date
--determine if the material should be expedited or postponed or asplanned

# steps of using codes

step1:run the data_interaction.py to get BIM_extended file

step2:run the data_merge.py to link building element (BIM_extended file) with construtcion look ahead schedule
