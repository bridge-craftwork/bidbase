# CardShark BidBase, 2026 (Nelson Ford)

Nelson Ford's current BidBase programs and rules, as he put them on his
web site in October 2026 for this project. The September 2018 SourceForge
release is in `../sourceforge-2018`.

- `source/`: the Visual Basic 6 source and build files (80 files). 77
  come from `http://www.aeyec.com/bidbase/Source/BidBase.zip` (uploaded
  2026-10-11); `CARDSH.ICO`, `MultiMon new.bas` and `cardshark icon.jpg`
  were in the same folder but not in the zip. File dates are Nelson's.
  - `BidBaseDB_Bidder.cls`: the bidding engine (the DLL), changed
    2026-09-28.
  - `BidBase Editor.frm`, `BidBase_Editor*.vbp`: the BidBase Editor, where
    the rules are written and tested.
  - `BidBase Bidding Practice.*`, `BidBase Bidding Practice Dealer.bas`:
    the practice program, which deals and bids hands.
  - `BidBaseFileViewer*`: the File Viewer for the documentation.
  - `BidSim.frm`, `BidBase_BidSim.vbp`: a bidding simulator (2019).
  - `BidBase Eval.bas`: hand evaluation; `dds.bas`, `dds.dll`: the
    double-dummy solver the engine calls.
  - The `.exe`, `.dll`, `.ocx` and `.lib` files are Nelson's builds and
    the run-time pieces the projects reference.
- `BidBase database.zip`: `BidBase.accdb` (178 MB, Access ACE12), the
  rules as of 2026-10-09 (`http://www.aeyec.com/bidbase/Source/BidBase.accdb`),
  zipped to stay under GitHub's file limit. `../data/2026` is its CSV
  export (`tools/export-db.sh`), which the Rust port reads.

The engine's changes since 2018 are in `../docs/ENGINE-2026.md`.

BidBase is Nelson's work. His web site (http://www.aeyec.com/bidbase/)
says: "All the databases, text files and programs making up BidBase are
free public domain", and BidBaseDB.DLL is free for use in any bridge
software. On 13 October 2026 he gave this project his specific
permission to publish and distribute these files, as they are or modified
(`../PERMISSION.md`).

The executables were not run in this project; build or run them on
Windows from Nelson's own copies.
