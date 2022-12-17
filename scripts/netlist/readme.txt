Is the allegro_netlist.py a two-step approach to generate the allegro netlist?

Through "Add Generator" button to add allegro_netlist.py in KiCAD's Export Netlist window to generate a KiCAD's XML format netlist;
Through the command line "python allegro_netlist.py flat_hierarchy.xml output_dir" to generate Allegro's netlist.

The script generates a netlist.txt (named the same as the xml) as well as a device file for every type of component used, which go into the devices subdirectory next to the netlist file.

If the component as a field named "AllegroFootprint", this one will pe picked-up as footprint for allegro.
