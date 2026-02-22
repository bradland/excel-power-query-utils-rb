# excel-power-query-utils-rb

Utilities for extracting, repacking, and refreshing Power Query (Data Mashup) content in Microsoft Excel (.xlsx) workbooks using Ruby.

**Highlights**
- **Extract**: Pull the internal Data Mashup (Power Query) binary out of an .xlsx and unpack its contents.
- **Repack**: Rebuild the Data Mashup from an unpacked directory and inject it back into an .xlsx.
- **Refresh (Windows only)**: Automate Excel to refresh Power Query connections and persist results.

**Requirements**
- Ruby (2.5+ recommended)
- Gems: `rubyzip` (for ZIP handling). Other used libraries (`base64`, `rexml`, `fileutils`, `optparse`) are part of the standard library.
- For `refresh_power_queries.rb`: Microsoft Excel on Windows (uses Win32 OLE automation).

Installation
1. Install Ruby (if not installed).
2. Install the required gem:

```bash
gem install rubyzip
```

Quick Usage

- Extract Data Mashup contents:

```bash
./extract_power_queries.rb -o out_dir --split workbook.xlsx
# or
ruby extract_power_queries.rb -o out_dir --split workbook.xlsx
```

- Repack an unpacked Data Mashup directory into a new Excel file:

```bash
./repack_power_queries.rb -s unpacked_dir -t template.xlsx -o output.xlsx
# or
ruby repack_power_queries.rb -s unpacked_dir -t template.xlsx -o output.xlsx
```

- Refresh Power Query connections (Windows + Excel required):

```powershell
.
# On Windows: run the refresher against the workbook
./refresh_power_queries.rb updated_report.xlsx
# or
ruby refresh_power_queries.rb updated_report.xlsx
```

Notes & Tips
- The extractor looks for the Base64 DataMashup element inside `customXml/item*.xml` and extracts the inner ZIP stream, placing files into the specified output directory. Use `--split` to separate individual `shared` queries from `Formulas/Section1.m`.
- The repacker re-creates the inner ZIP, preserves the Microsoft MS-QDEFF header from a template workbook, encodes the combined binary as Base64, and replaces the `<DataMashup>` XML content.
- The refresher requires Windows because it uses Win32 OLE to control Excel and force synchronous refreshes.

Documentation
- To generate the API documentation locally, use `rdoc` and output into `doc/`:

```bash
# Install rdoc if needed
gem install rdoc

# Generate documentation into the repo-local `doc/` directory
rdoc -o doc .
```

- After generation open the index in your browser:

```bash
open doc/index.html
```

Files
- Extractor: [extract_power_queries.rb](extract_power_queries.rb)
- Repacker: [repack_power_queries.rb](repack_power_queries.rb)
- Refresher: [refresh_power_queries.rb](refresh_power_queries.rb)

License
- See the repository `LICENSE` file for licensing information.

Contributing
- Bug reports, suggestions, and PRs are welcome. Please open issues or pull requests in the repository.

Contact
- For questions about usage or edge cases, open an issue in this repository.
