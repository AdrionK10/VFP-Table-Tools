Migrating SQL Tables to Visual FoxPro Using Visual FoxPro Table Tools Program (VFPTT)


Follow the steps below to efficiently move your SQL tables to VFP using the VFPTT:
1. Develop a data access entity, services, controller, and entity mappings for the project.

2. Launch the VFPTT and select your local connection.

3. Create a new table, ensuring that the columns correspond to the entity types defined in Step 1. Keep column names within a 10-character limit and maintain consistency with the entity naming conventions.

4. The newly created table will be added to the Vista database! If you have data for this table to import, proceed to the next steps.

5. Open Table Tools (vista.dbc) from the VFPTT main menu (Option 2).

6. Choose the newly created table. You will now have access to various options, such as inspecting the schema, exporting to CSV, importing, mirroring, disassociating, and deleting.

� Inspect: Displays table information, including row count and column details.

� Export to CSV: Exports all table data to a CSV file at the specified location (used for import and mirror).

� Import: Imports data from the CSV file. Modify the exported CSV file to update, delete, or create data in the table. Rows omitted here will not be removed from the database. This tool can create or update a single record based on the CSV file content.

� Mirror: Imports data from the CSV file. Modify the exported CSV file to update, delete, or create data in the table. Rows omitted here will be removed from the database; mirroring an empty CSV file will delete all rows in the database.

� Delete: Removes the table from the vista.dbc database.

Typically, you would export the data, paste it into the exported CSV file, and then perform an import.

Additional Options:
� If you have recently retrieved data from the Thrive server, you may need to rebuild the table. You will notice this when opening Table Tools (vista.dbc) from the main menu (Option 2) and the desired table is not listed. To make the table appear, open Table Tools (other) from the main menu (Option 3). You should find the table there. Select the table and choose Option 1 to rebuild the table.
