# The Bidding Practice program

A behaviour spec of Nelson Ford's "CardShark BidBase Practice Dealer"
(`aeyec-2026/source/`, form `PlayScr`, caption `CardShark BidBase Practice
Dealer     ver.2025.07.31`), written so it can be ported to Rust/egui
without reading the VB again. Only the source was read; no `.exe` or
`.dll` was run. The compiled `BidBase Bidding Practice.exe` is from June
2022 and older than this source (October 2026).

Line references:

| Prefix | File |
|---|---|
| `frm:` | `BidBase Bidding Practice.frm` (layout 1-1687, code 1688-8417) |
| `dealer:` | `BidBase Bidding Practice Dealer.bas` (module `DealerRoutines`) |
| `dds:` | `dds.bas` |
| `eval:` | `BidBase Eval.bas` (module `BBEval`) |
| `cls:` | `BidBaseDB_Bidder.cls` (the engine, 2026) |

Sizes in the `.frm` are twips (15 per pixel at 96 dpi). At run time the
form's `ScaleMode` is pixels (frm:6989), so code that moves controls on
the form uses pixels; controls inside frames stay in twips. Pixel
positions below are twips / 15.

Seats are numbered **1 = South, 2 = West, 3 = North, 4 = East**
everywhere in the program (clockwise from South). `int_Player(k)` is the
seat of the k-th caller of a round, counted from the dealer (frm:6858).
Two other numberings turn up in the DDA code (see section 5).

Several features are unfinished in this source. A `Stop` statement
halts the program in the VB IDE and ends it in a compiled exe; these sit on
live paths at frm:2239, 2282, 2322, 2339, 4115, 4134 and 4965. Debug
message boxes ("Stop MakeBid 1" and so on) are at frm:5347, 5844, 6290,
6316, 6732, 6745, 6796, 8161 and dealer:423. Where this spec says
"unfinished", it means one of these is reached.

## 1. Screen layout

The form is 664 x 576 pixels (client area), fixed border, no maximize
button, background green `&H00C000` (RGB 0,192,0). Default font MS Sans
Serif 8.25 bold. It starts always-on-top (frm:3382 toggles Stay on Top on).
The window position is restored from the config file.

```
 +--------+----------------- North hand -----------------+--------+
 |        |                                              |        |
 |  West  |  +-- bid box --+  +-- auction --+            |  East  |
 |  hand  |  | 1C 1D 1H 1S 1N | S  W  N  E |             |  hand  |
 | (fan   |  | ...        7N | .. .. .. .. |             | (fan   |
 |  down) |  | Pass Dbl Rdbl BID BOX      |             |  down) |
 |        |  +---- t_Disclose (disclosure text) ---+     |        |
 |        |                                              |        |
 |        +----------------- South hand -----------------+        |
 +---------------------------------------------------------------+
 | Vulner. none  | Deal  #31118      |  BB Bid  | Enter | Bid w/  |
 | Dealer  W Keep| < Repeat >        |  Edit Bid| Hands | Analysis|
 +---------------------------------------------------------------+
 | cmb_DDA: DDA table, 5 lines (11 with averaging), Courier bold   |
 +---------------------------------------------------------------+
 | cmb_FindDeals: convention auction specs (only when enlarged)    |
```

### Hands (`p_Hand(1..4)`)

| Control | Seat | Pixels (x, y, w, h) | Fan |
|---|---|---|---|
| `p_Hand(1)` | South | 169, 330, 328, 98 | horizontal |
| `p_Hand(2)` | West | 2, 4, 77, 422 | vertical |
| `p_Hand(3)` | North | 168, 0, 325, 98 | horizontal |
| `p_Hand(4)` | East | 591, 4, 77, 424 | vertical |

Picture boxes, green, no border (frm:350-629). Cards are drawn into them by
`DisplayHand` (dealer:178):

- Cards come from `Qcard32.dll`: face 71 x 96 pixels (`CARDWID`,
  `CARDHT`, dealer:59). Calls used: `InitializeDeck(hwnd)` once at start
  (frm:3336; failure shows "Problem loading QCard32.DLL. ..." and
  unloads), `DrawCard(hwnd, card 1-52, x, y)`, `DrawBack(hwnd, 3, x, y)`
  (back design 3, red), `GetCardSuit(card)` (1 = clubs .. 4 = spades) and
  `GetCardValue(card)` (1 = ace, 2-13 = two to king). Everything else
  declared at dealer:87-146 is unused.
- Card numbers: 1-13 clubs, 14-26 diamonds, 27-39 hearts, 40-52 spades;
  within a suit 1 = ace, then 2..13 = two..king. So `card = (suit-1)*13 +
  rank`.
- Order: spades, hearts, diamonds, clubs; in each suit ace first, then
  king down to two. Each card is drawn over the previous one.
- South and North: each card 21 pixels right of the previous. West and
  East: each card 27 pixels below the previous. The first card is at
  (0, 0). A full hand fills 323 x 96 (N/S) or 71 x 420 (W/E).
- Hearts/diamonds divider: the ace of diamonds, or a diamond drawn right
  after a heart or spade, is drawn three times at 1-2 pixel offsets so a
  thin edge separates the two red suits (dealer:226-275). A port can
  draw a small gap instead.
- Hidden hands: when Hide Hands is on, West, North and East are drawn as
  card backs in the same fan positions; South is always face up
  (dealer:222, 252).
- `DisplayHand` does not clear the picture; callers call `p_Hand.Cls`
  first where it matters. In egui just redraw.

### Bid box (`g_BidBox`)

MSFlexGrid at pixels 170, 103 (171 x 138), 8 rows x 5 columns, no fixed
rows, yellow (RGB 255,255,0), green grid lines (frm:496). Filled by
`SetUpBiddingGrid` (frm:1950): rows 0-6 are levels 1-7, columns C, D, H,
S, N; cell text is level + strain + a space ("1C "). Row 7: "Pass", " Dbl",
"Rdbl", " BID", "BOX". The last two cells are a label; a picture
(`p_BidBoxLbl`, frm:485) sits over them and clicks on them are ignored
(frm:3482). Column width 500 twips, row 7 height 300 twips.

A click gives the call `Trim(text)` ("1C", "7N"), or for row 7 the first
letter: "P", "D", "R" (frm:3481-3487). What happens next: section 2.

### Auction grid (`g_Bidding`)

MSFlexGrid at pixels 344, 103 (147 x 138), 7 rows x 4 columns, light
green (RGB 128,255,128), **disabled** (`Enabled = False`, frm:1477), so it
cannot be clicked and has no handler. Row 0 is a fixed header "S", "W",
"N", "E" (frm:1928; column widths 530/530/530/535 twips, header row 260).
Bold 9.75 pt.

`DisplayBid` (frm:6752) clears the grid and writes the calls row by row:
the dealer's column on row 1, then left to right, wrapping to the next
row after East. Calls are shown as BidBase writes them: "1C", "1N", "P",
"D", "R". A call carrying an ace-asking code ("4N@30", section 3) shows
only its first two characters. There is no per-cell colour. Six rows hold
at most 24 calls less the dealer offset; a longer auction shows "Stop
DisplasyhBid 1" (frm:6796). The second half of `DisplayBid` (frm:6803-6844)
writes the last call again using a W-N-E-S column layout; it is only
reached if the first loop never meets an empty call, so it is dead in
practice.

### Disclosure box (`t_Disclose`)

Multi-line text box with a vertical scroll bar, pixels 168, 242 (324 x
85), dark green text (RGB 0,64,0) (frm:515). It shows the engine's
disclosure for the last call, messages ("*** Your bid is the same as BB's.
***"), the engine's `ErrMsg`, PBN players, or PBN commentary. Double-click:
section 2.

### Bottom panel (`p_Bottom`)

Green picture box at pixels 2, 427 (659 x 54) (frm:1287), plus three
controls placed on the form over it:

| Control | Caption | Where | Notes |
|---|---|---|---|
| `b_Vul` | `&Vulner.` | left, top | label `L_Vul` to its right: "none" at start, then "", "Unfav.", "Fav.", "Both" |
| `b_Dealer` | `Dea&ler` | left, bottom | label `L_Dealer` (S/W/N/E), then checkbox `ck_KeepDealer` "Keep" |
| `b_Deal` | `&Deal` | middle, top | label `L_Deal` to its right: deal number, "User", "User - name", or "round+board - file" for PBN |
| `b_DealBackFwd(0)` | `<` | middle, bottom | |
| `b_RestartDeal` | `&Repeat` | between the arrows | |
| `b_DealBackFwd(1)` | `>` | | |
| `b_FinishedAuctions` | `Finished Auctions` | middle, bottom | hidden except while editing convention specs |
| `b_Ok` | `&BB Bid` | right, top | the default button (Enter) |
| `b_Edit` | `&Edit Bid` | right, bottom | |
| `b_EnterHands` | `Enter Hands` | pixels 515, 430 | in frame `f_NSEW` |
| `b_ShowAnalysis` | `Bid w/ &Analysis` | pixels 588, 430 | |

`lblTime`/`lblBattTime` ("Battery Left:") and `ck_FindBids` ("Find/practice
conventions") are hidden and dead; the battery code is commented out
(frm:7016-7060).

Vulnerability labels are from South's side: "Unfav." means North-South
vulnerable only, "Fav." East-West only, "" none.

### DDA table (`cmb_DDA`)

A locked "simple combo" (a list with a text line on top), pixels 2, 479
(658 x 90), Courier 9.75 bold (frm:466). Its text line reads `DDA =
Maximum tricks that can be taken by each player in each suit/NT.` Filled by
`ShowDDA` (frm:7161) with five lines, in this order:

```
DDA:C  D  H  S  N | HCP | LoTT | LTC: Calculate Tricks
N: <dda N> | <HCP N 3 wide> |      | <ltc N>   Pairs Can Take:
S: <dda S> |  <HCP S 2 wide> | <LoTT NS 3 wide>  | <ltc S><caption NS>
E: <dda E> |  <HCP E 2 wide> |      | <ltc E>
W: <dda W> |  <HCP W 2 wide> | <LoTT EW 3 wide>  | <ltc W><caption EW>
```

Exact pieces (frm:7238-7521):

- `<dda X>`: double-dummy tricks for clubs, diamonds, hearts, spades,
  notrump; the first number right-aligned in 2 characters, the others in
  3, then trimmed and given back one leading space if 13 characters long.
  So always 14 characters, e.g. ` 7  8 10  9  6` or `10  8 10  9  6`.
- HCP: A 4, K 3, Q 2, J 1, counted from the hand string. North's is
  `Right("  " & n, 3)`; the others `" " & Right(" " & n, 2)`.
- LoTT ("Law of Total Tricks"): the longest combined suit of the pair,
  ties to the higher suit (frm:7333-7342), `Right("  " & n, 3)`, shown on
  the South line for NS and the West line for EW.
- LTC: losing-trick count per hand (frm:7361-7441): void 0; singleton 1
  unless the ace; doubleton AK 0, AQ 0.5, Ax 1, Kx 1, else 2; three or
  more cards: AKQ 0, AKJ 0.5, AK 1, AQJ 1, AQ 1.5, A 2, KQ 1, K 2, QJ or
  QT 2, Q 2.5, else 3. Printed as VB prints a Single ("7", "7.5", "10"),
  padded: one character becomes `" 7 "`, two characters get a trailing
  space.
- Captions: on the South line `-> 24 - <ltc N> - <ltc S> = <24 - both>`,
  on the West line `-> 24 - <ltc E> - <ltc W> = <...>`; replaced by `->
  No 8-card fit.` when the pair has no 8-card fit and no "self-sufficient"
  suit (length + honours > 9). For EW the self-sufficient test adds HCP
  instead of honours (frm:7353-7354, a slip).

A worked line (South 12 HCP, 7 losers; North 6 losers; NS longest fit
9): `S:  7  8 10  9  6 |  12 |   9  |  7 -> 24 - 6 - 7 = 11`. With a
three-character LTC such as "7.5" there is no space before "->".

With averaging on (Options, 11-Deal Avg. DDA) the form grows to 10510
twips high, the list to 180 pixels, and six more lines follow; section 5.

`ShowDDA` also writes `DDAresults.txt` (section 7) and computes the
"optimum" contracts passed to the engine (section 5).

### Message frames

All are drawn over the form; the VB keeps them modal by looping on
`DoEvents` until they hide.

- `f_Msg` (frm:40): the program's own message box, used through
  `MyMsgBox(text, numLines, response, caption)` (frm:8247). Red title
  strip ("  CardShark BidBase" unless a caption is given), white body,
  message in 12 pt bold. `response` "ok" shows OK; "YN" shows Yes and No;
  "YNC" adds Cancel; anything else shows a text box with that default and
  an Ok button. Returns "yes"/"no"/"cancel", the typed text, or "" for OK.
  Width comes from `TextPixelWidth` (frm:8370, a per-character width
  table); callers set its top (200, 360, 480 or 538 pixels). While it is
  open, `t_Calculating` (a green text box "Calculating DDA" over the
  bottom panel, frm:1505) covers the buttons.
- `f_Inbox` (frm:199): small prompt with a 12 pt centred text box and a
  Done button. Used for the deal number ("Enter a deal number or 'user' to
  enter hands:", dealer:450) and the dealer ("Enter dealer" / "(N S E W)",
  frm:1996). For the dealer prompt, typing S, W, N or E closes it at once
  (frm:5705).
- `f_BillBoard` (frm:408): light green notice with OK; only used by Find
  matching deals (frm:4085).
- `f_ClearDeal` (frm:278): white box "Clear the existing deal and select
  new cards?" with Y and N; shown by Enter Hands.
- Plain `MsgBox` and `InputBox` are also used in many places; their texts
  are quoted where they matter below.

### Hand entry controls (hidden until Enter Hands)

- `t_EnterHand(1..4)` (frm:364-407): one text box per seat, 126 x 19
  pixels: South at 266, 303; West 70, 196; North 266, 124; East 462, 196.
  The seat being filled is yellow (RGB 255,255,0), the others pale green
  (RGB 204,255,204).
- `f_CardDeck` (frm:653): the card picker, pixels 68, 94 (521 x 238), 52
  picture boxes `Card(1..52)` of 40 x 57 pixels, so each shows only the top
  left corner of a card. Rows from the top: spades, hearts, diamonds,
  clubs; columns from the left: A, K, Q, J, T, 9 .. 2. A picked card is
  covered by a card back.
- `b_DealRest` "Deal the rest" (80, 337), `b_CancelHands` "Cancel" (505,
  337), and the help label `L_EnterHandsHelp` (12 pt): "Click on N, S, E,
  or W, then / click on cards for that hand. / or enter hands in text
  boxes." (dealer:317).

### Other hidden controls

- `cmb_FileList` (frm:171): list of convention files.
- `cmb_FindDeals` (frm:636): convention auction specs, below the DDA list;
  the form grows to 10770-11140 twips to show it.
- `cmb_PBNBoards` (frm:31): list of boards in a loaded PBN file.
- `Timer1` (frm:630): 200 ms one-shot used to run a job after the current
  handler returns (frm:7063): "Cancel entering hands", "Get next bid",
  "Restart same deal", "Go to next deal".
- `CmDia`: the common file dialog. Every "open" in the program uses
  `ShowSave`, a Save dialog.

## 2. User actions

### Start-up (`Form_Load`, frm:3292)

1. Paths: everything is relative to the program folder (`App.Path`); the
   database is `App.Path\BidBase.accdb`. If it is missing, a message and
   unload. A second copy of the program refuses to start ("Program is
   already running.").
2. Read `Last Deal Number.txt` (deal number minus 1), then `LoadCfg`
   (frm:6931), which reads `BBPractice.cfg` and overrides the deal number
   (section 7). `LoadCfg` also calls `Randomize`, shows the form at the
   saved position, runs `mHide_Click` (see Hide Hands), and creates the
   engine object `NewBid`.
3. Fill the bid box, deal (the number is incremented, so the last deal
   of the previous session comes back), draw the hands, `ShowDDA`.
4. `InitiateBidding` (frm:5755): dealer from `L_Dealer` (from the
   config), seat order, clear the auction, and the dealer's first call is
   made at once (`BidNextHand` -> `MakeBid`).
5. Stay on Top is switched on.

### Deal (`b_Deal_Click`, frm:2082), Alt+D

- If a user deal or a PBN event is showing: "This will end viewing
  <L_Deal> and return to random deals. Continue?" (Yes/No). Yes goes back
  to the random deal that was showing before (`lng_CurrentDeal`), arrows
  re-enabled, deal redrawn, number saved to `Last Deal Number.txt`. (The
  VB runs `FinishDealSetup`, hence `ShowDDA`, before `Deal`, so the DDA
  is for the old hands, frm:2104-2105.)
- Otherwise it prompts for a deal number with `f_Inbox`, default the
  current number (dealer:444). Above 9,999,999: "Deal number cannot be
  greater than 9,999,999." and ask again. The deal with that number is
  dealt and shown, the auction is cleared and the DDA computed
  (`FinishDealSetup`, frm:2140). The dealer does not change.
- No call is made; the first BB Bid click makes the dealer's call.
- Typing "user" was meant to start Enter Hands, but `ans` is never
  assigned (dealer:460), so "user" and 0 both just set an abort flag.
  `Rnd(-0)` does not reseed, so "deal 0" is not reproducible. A port
  should accept 1..9,999,999 only.

### < and > (`b_DealBackFwd_Click`, frm:2348)

- PBN event showing: previous/next board in the PBN list (wrapping), see
  section 7.
- Practising a convention: previous/next saved entry of the `.Deals` file
  ("This was the first saved deal." / at the end "This was the last saved
  deal." and practice stops).
- Otherwise: deal number - 1 or + 1. With Keep unchecked, > moves the
  dealer one seat clockwise (S, W, N, E, S...). < sets the dealer two
  seats back (comment: "It will be increased by 1", but nothing does),
  or to South when that goes below 1; so < does not restore the previous
  deal's dealer. A port should step the dealer back by one.
- Then: deal, draw, `FinishDealSetup` (clear auction, DDA), make the
  dealer's first call at once, write `Last Deal Number.txt`.
- The vulnerability never changes with the deal number.

### Repeat (`b_RestartDeal_Click`, frm:2917), Alt+R

Same deal, same dealer (taken from `L_Dealer` if it was changed), same
vulnerability. Clears the auction and redraws the hands; the DDA is not
recomputed. No call is made until BB Bid.

### Vulner. (`b_Vul_Click`, frm:2978), Alt+V

Cycles `L_Vul`: "" -> "Unfav." -> "Fav." -> "Both" -> "". (The initial
caption "none" goes to "" on the first click.) The auction is not
restarted; the help file says to click Repeat.

### Dealer (`b_Dealer_Click`, frm:2453), Alt+L

Unchecks Keep, asks for the dealer with `f_Inbox` (S/W/N/E; a key press
closes it; anything else becomes South), shows Keep, then Repeat.

### Keep (`ck_KeepDealer`)

When checked, > and < keep the dealer and the Dealer prompt is skipped
(`GetDealer`, frm:1992).

### BB Bid (`b_Ok_Click`, frm:2871), Alt+B or Enter

1. Writes the window position to `DealerPosition.txt` (section 7).
2. `ProcessBid` (frm:5927): if the auction already ends with three passes
   after a call, or four passes, it shows "Three consecutive passes." /
   "Bidding's over." and stops. Otherwise the engine makes the next call
   for the seat whose turn it is (section 3).

The captions "Deal" and "&Continue" that `b_Ok_Click` also handles are
never set (frm:6053 needs a pass count that stays 0), so they are dead.

### Edit Bid (`b_Edit_Click`, frm:2648), Alt+E

Hands the last call to the BidBase Editor and takes it back:

1. Writes `App.Path\EditPractice.txt`, three lines: the hand of the
   player who made the last call; the calls before it joined with "-"
   (with any "?" removed); and `Deal #<L_Deal>-<L_Dealer> ID=<rule id>
   BON=<bid order number>` (for user and PBN deals only `Deal
   #<L_Deal>-<L_Dealer>`). The Editor imports this file (its Import
   command reads hand, prior calls, source, `ID=`, `BON=`).
2. Removes the last call from the auction and the grid; clears the
   disclosure.
3. The next BB Bid re-bids that seat, so an entry edited in the Editor
   takes effect.

### Bid w/ Analysis (`b_ShowAnalysis_Click`, frm:2931), Alt+A

- PBN event showing: puts the board's commentary (`{...}` block) in the
  disclosure box.
- Otherwise: InputBox "Enter bid to watch for or ALL:". Upper-cased; any
  "all" becomes "ALL"; a value starting with a level n gets " nx"
  appended ("2S" -> "2S 2x", because a rule bidding "2x" may be 2S). The
  current call is cleared and the seat re-bid with `BidToAnalyze` set; the
  engine then shows its analysis form listing, for every rule with a
  watched bid, the test it failed. Afterwards the caption becomes
  "Rebid w/Analysis".
- The re-bid goes through `BidNextHand` with the rebid flag, which also
  clears call `int_CallNum - 1` (frm:5876-5878) after frm:2966 has
  cleared call `int_CallNum`; the indices look off by one. A port should
  re-run the lookup for the last call with tracing on (the port's
  `Outcome.tried`, the CLI's `--why`).

### Clicking the bid box (`g_BidBox_Click`, frm:3475)

The call clicked is `str_MyBid`. Two cases:

1. **The last call shown differs from the click**: "Do you want to add
   your bid of <mine>" / "to BidBase to supersede BB's bid of <BB's>?"
   (Yes/No). Yes: the engine inserts a new rule for this hand and auction
   with the user's call (`AddBidToBB`, frm:3562: `Job = "Save"`, see
   section 3), the last call is replaced by the user's, and the handler
   ends. No: go on to case 2.
2. Otherwise (or after No): the next empty call slot is found; the engine
   makes its call for that seat; if it differs from the click:
   "Did you want to change BidBase's bid of <BB's>" / " to your bid of
   <mine> in the database?" / "(Click No to keep your bid for now.)"
   (Yes/No/Cancel). Yes runs Edit Bid (export to the Editor, call
   removed). No keeps the user's call in the grid. Cancel leaves BB's call.
   If they agree, the disclosure box gets:

   ```
            *** Your bid is the same as BB's. ***

               Click a bid for the next player or
               click "BB Bid" for BB to do it
               or click right arrow for next deal.

   <disclosure>
   ```

Problems to avoid in a port:

- After "No" in case 2 the grid shows the user's call but the auction
  string sent to the engine keeps BB's call (frm:3548 changes `str_Call`
  only; `MakeBid` already appended BB's call to `str_AllCalls`).
- Case 1 removes every "-<BB's call>" from the auction string
  (`Replace`, frm:3498), not just the last one.
- On a fresh deal (no call yet) the call slot is 0 and the engine is given
  `int_Player(4)`'s hand for the dealer's call (frm:3526-3529).
- The "You bid ... and BidBase bid ..." banner in `MakeBid`
  (frm:6321-6344) always says "same" because it compares the auction's tail
  with the call just appended.

Intended behaviour, from the 2018 help page
(`Notes/CardShark_BidBase_Practice_Program.htm` in the SourceForge File
Viewer zip): with Auto-Bid off you bid first (all seats if hands are
shown, only South if hidden), BidBase then says what it would have bid;
clicking a different call right after BidBase bid is a way to replace
that call and send the hand to the Editor.

### Clicking the auction grid

Nothing: the grid is disabled.

### Double-clicking the disclosure (`t_Disclose_DblClick`, frm:5719)

Takes the first line of the disclosure; if it contains ":", the part
before the colon with `/\.?^*<>|` and spaces replaced by "_". Opens
`Notes\<name>.htm` in the File Viewer, or "No file for "<name>"". The VB
launches `BidBase FileViewer.exe` (no space) while every other place uses
`BidBase File Viewer.exe`, so this fails as written.

### Enter Hands (`b_EnterHands_Click`, frm:2702; also Options, Enter Hands)

1. Disables Deal, Repeat, Bid w/ Analysis, Enter Hands; shows the card
   picker with all cards face down, then `f_ClearDeal` "Clear the
   existing deal and select new cards?".
2. Y: all hands cleared, the picker shows all 52 faces (`LayOutDeck`,
   dealer:289), South's box is the target (yellow). N: the current hands
   are put in the four text boxes for editing.
3. The four text boxes, "Deal the rest" and "Cancel" appear.

While entering:

- **Click a seat's text box or hand** to make it the target (yellow).
- **Click a card in the picker** (`Card_Click`, frm:2992): "Hand already
  has 13 cards." if full; "Click on N, S, E or W first." if no target;
  ignored if already dealt. Otherwise the card goes to the target hand,
  its picker cell turns into a back, the text box and the hand are
  redrawn. When the hand reaches 13 its box turns green and the next
  incomplete seat becomes the target, with the message "Click an input
  box / to deal to or / click Cancel or / click Deal The Rest."; when all
  four are complete the button caption becomes "DONE".
- **Click a card in a hand picture** (`p_Hand_MouseDown`, frm:5307):
  returns it to the deck. Hit test: N/S 300 twips per card across, W/E 400
  twips down, while drawing uses 21 and 27 pixels (315 and 405 twips), so
  clicks near the end of a long hand pick the wrong card. A port should
  hit-test against the drawn positions.
- **Type in a text box** (frm:5457-5617): format `AKQ-JT9-876-5432`
  (spades-hearts-diamonds-clubs). Space becomes "-"; a, k, q, j, t are
  upper-cased; only AKQJT, 2-9, "-", Backspace and Ctrl+C/V/X are
  accepted; a card ranked at or above the one before the cursor gives
  "'<c>' is out of order."; deleting a selection keeps its dashes. Each
  change re-parses the box: "Too many dashes.", "<suit letter>:<card> has
  already been dealt.".
- **Deal the rest** (`b_DealRest_Click`, frm:2460): gives the undealt
  cards to the seats whose box is empty. The VB assigns them before it
  shuffles (frm:2569-2598), so the first empty seat gets the 13 highest
  remaining cards; seats with partial hands are not completed. A port
  should shuffle the remaining cards and complete every short hand.
  The caption then becomes "DONE".
- **DONE**: hides the picker, `L_Deal` = "User"; "Save deal?" (Yes/No);
  Yes opens a file dialog in `UserEnteredHands\` (`*.txt`), "Overwrite
  existing file with that name?" (Yes/No/Cancel; No saves as `temp.txt`
  in the current directory); format in section 7. Then "Enter
  Vulnerability and Dealer and click Repeat." and the DDA.
- **Cancel** (`b_CancelHands_Click`, frm:2022): restores the buttons and
  redeals what was showing: a user deal is reloaded from its file (the
  name after "User - "), a numbered deal is dealt again.

`mEnterHands_Click` (frm:3727) calls the same handler and exits; the rest
of it (InputBox entry per seat, `ValidateHand`, `FillInTheDeal`, a
"source" prompt and a prior-calls prompt) is dead, as are `AddHandToDeal`
(frm:1861) and `PutCallsIntoArray` (frm:4008, which never reads its
input).

### File menu

| Item | Key | Size | What it does |
|---|---|---|---|
| Add Bids not in BidBase | | | hidden and disabled; no handler. The feature lives in the bid box (case 1 above). |
| Save Deal | Ctrl+S | small | appends the deal to `Notes\SavedDeals.txt` and opens it in Notepad (format in section 7). For review, not reloading. |
| Load User Deal | Ctrl+L | small | "The loaded deal will replace the current deal. Continue?"; file dialog in `UserEnteredHands\`; `L_Deal` = "User - <name>"; "Enter dealer and vulnerability / then click Repeat, if necessary."; arrows disabled; DDA. |
| Load PBN Event | Ctrl+P | medium | section 7. |
| Load PBN Deal # | Ctrl+D | small | enabled after a PBN load; shows the board list; double-click a board to show it. |
| Run Editor | | external | runs `BidBase Editor.exe`. (It first tries `AppActivate "BidBase Editor.exe"`, which always fails, and checks for `BidBase Bidding Practice.exe` instead of the Editor, frm:5291-5298.) |
| Quit | Ctrl+Q | small | saves `BBPractice.cfg` and ends. |

### Options menu

| Item | Key | What it does |
|---|---|---|
| Auto-Bid | | toggles `bln_AutoBid` (checked = the computer bids every seat). Unchecking enables Hide Hands; checking clears it (frm:3583). In this source its only effect is on the user-bids-first path of `BidNextHand` (frm:5889) and the banner in `MakeBid`, both reached only from the bid box. |
| Hide Hands | Ctrl+H | toggles drawing W, N, E face down (frm:4208). |
| Stay on Top | Ctrl+T | toggles always-on-top. |
| Enter Hands | | as the button. |
| Redisplay Hands | F12 | redraws the four hands. |
| Mix Up Cards | Ctrl+M | InputBox "Enter the hands (N,S,E,W) to scramble."; 2 or 3 seats ("You must specify at least 2 hands to Scramble." / "You cannot Scramble more than 3 hands. To Scramble all 4, just click on Deal."). The cards of those seats are pooled, shuffled (same loop as the deal, section 4, without reseeding) and dealt back to the same seats; redraw, DDA, Repeat. Enables Remix. |
| Remix Same Hands | Ctrl+R | shuffles the same seats again ("No record of hands selected. Please enter again." if none). |
| 11-Deal Avg. DDA | | toggles averaged DDA (section 5); medium-sized feature. While checked, every new deal recomputes it. Unchecking removes the extra lines. |
| 101-Deal Average | Ctrl+A | enabled while the 11-deal item is checked; one run with 110 layouts per pair (the label says 101). |
| Rotate Hands W-S / N-S / E-S | Ctrl+W / Ctrl+N / Ctrl+E | moves the West (North, East) hand to South, the others following counter-clockwise; the dealer moves with the cards (frm:5033). The menu index of E-S is 23, not 3: 23 rotations give the same hands as 3, but the dealer arithmetic (`int_Dealer - 23`, then one `+ 4`) goes negative. The VB rotates `int_Cards` wrongly (frm:5056-5064 copies every block into slots 1-13), so a later Mix or averaged DDA works on stale cards, and the DDA table is not refreshed. A port should rotate one seat array and recompute. |

### Conventions menu (convention practice): large, unfinished

Meant to find deals whose auction starts with given sequences, save them,
and quiz the user on the next call. In this source "Find matching
deals" and "Practice convention" both reach a `Stop`
(frm:2282/4115/4134 and frm:4965), so they end the compiled program. The
design, from the code and comments:

1. **How conv. practice works**: opens
   `Notes\CardShark_BidBase_Practice_Program.htm#C1` in the File Viewer.
2. **Create/pick convention file** (`mCreateConv_Click`, frm:3619;
   toggles). Refuses while Find or Practice is checked. `OpenPracticeFile`
   (frm:4741) makes `PracticeDeals\` and `PracticeDeals\Deals\` and lists
   the files of `PracticeDeals\` under four header lines: "Enter a new
   convention above.", "     Click this line to cancel.", "Click a line
   below to select.", "------...". Clicking a line selects it (the header
   lines 0, 2, 3 are taken as names too; there is no handler for typing a
   new name). Then the spec list `cmb_FindDeals` appears under the DDA
   (form height 10770 twips) with the text "Enter an auction here or click
   an auction to edit or click Finished.", loaded from
   `PracticeDeals\<name>` if it exists, and the "Finished Auctions" button.
   - Typing a spec (`cmb_FindDeals_KeyPress`, frm:7102): Enter adds the
     line. Allowed: C D H S N P D R, 1-7, "-", "|", ",", space after a
     comma (any other space becomes "-"), "x" (forced lower case), "A". A
     comma is allowed only after a strain or P/D/R and becomes ", ". A
     strain letter must follow a level digit (D alone is a double).
   - Double-click a spec: "Delete?". Blank lines are removed on focus.
   - Finished Auctions (frm:2814): rewrites `PracticeDeals\<name>`, one
     spec per line; enables Find, and Practice/Update if
     `PracticeDeals\Deals\<name>.Deals` is over 10 bytes; then "Ready to
     find matching deals?" Yes runs Find.
3. **Find matching deals** (`mFindDeals_Click`, frm:4064). Billboard:
   "<name> will have 100 deals added to it. More deals can be added in
   the future by clicking this menu function. To stop adding deals at any
   time, click the menu again. Click Ok below to continue." Creates the
   `.Deals` file with an empty header record, starts at deal 1, dealer
   South. Intended loop (`ProcessFindBids`, frm:6360, driven by the
   timer): for each deal number and each dealer S, W, N, E, let the
   engine bid call by call; after each call compare the auction so far with
   the specs in order:
   - a spec matches by prefix on the "-"-joined calls;
   - "|" is ignored for matching but marks that the call after it is
     hidden: it is blanked and is what the user must find;
   - "x<call>" at a position means any call except that one;
   - "A" is any call;
   - every spec without "|" gets "-A" appended (frm:6453, "just for making
     quizzes"), so a match runs one call past the spec;
   - on a prefix match keep bidding; on a full match save the deal
     (number and dealer), append a line to `<name>.Txt` and a quiz item to
     `Notes\Practice_<name>.txt`; on no match try the next spec, then the
     next dealer, then the next deal;
   - only 2 in 10 matches of "1N-P-2C" are kept (the `stayman` counter,
     frm:6534);
   - after 100 entries: "100 new entries have been saved. Click Practice
     in the Conventions menu to use them."
   `FindPracticeDeals` (frm:2276), a rewrite of this loop, is a stub.
4. **Practice convention** (`mPractice_Click`, frm:4860; toggles). Picks
   a file (none yet: goes to Create). Reads the header; if the last entry
   used is past 2, InputBox "What file entry number do you want to start
   with? (Press Enter for the last entry used as shown.)"; wraps to the
   first entry at the end; deals that entry's deal number with its saved
   dealer. Intended next (timer job "Restart same deal"): the engine bids
   the spec's calls; when the spec is complete the hands are rotated so
   South is next to bid (`mRotate` by the last bidder, frm:6649-6660),
   the DDA is shown, and the user bids South. < and > step through the
   saved entries. If the engine's calls no longer match any spec: "Bidding
   for deal #<n> does not match the specified auctions. Most likely bid
   specifications in the database have changed. Skip to the next saved
   deal? (No to exit.)" and the entry is deleted (`DelPracticeDeal`
   copies the last record over it).
5. **Update convention file** (`mUpdateConventionFiles_Click`, frm:4823;
   toggles). "Get a 2nd bid for practicing?" (Yes/No/Cancel), then
   Practice in a mode that rewrites the quiz entry for each saved deal and
   moves on without stopping. With a 2nd bid the quiz entry also shows the
   next three calls, partner's hand and BidBase's second call.
6. **Show index to conventions**: "A directory of convention docs /
   arranged in convention card order / will be displayed next. Click on /
   a convention to see its doc." then `Notes\Index.htm` in the File
   Viewer.

The 2018 help page describes an older, simpler form of this: a "Start
with" box taking specs like "1H-P, 1S-P", with North as a kept dealer.

### Help menu

- Help (F1): `Notes\CardShark_BidBase_Practice_Program.htm` in `BidBase
  File Viewer.exe`, or "The Notes directory is not present. Download the
  BidBase Notes.Zip file or go to www.aeyec.com/BidBase."
- Tutorial (F4) and Version: hidden and disabled.

## 3. The bidding loop

### State

- `str_Call(1..50)`: the calls, as BidBase tokens ("P", "D", "R", "1C",
  "4N@30").
- `str_AllCalls`: the same joined with "-", opening passes included
  ("P-1C-P-1H"). This is what the engine gets.
- `int_CallNum`: the slot of the last call made (0 before any).
- `int_Turn` 1-4: position in the round; the seat is
  `int_Player(int_Turn)`.

### Who bids

Each BB Bid click makes exactly one call, for the next seat, by the
engine; after > (and at start-up) the dealer's call is made at once. So
in practice the computer bids all four seats and the user steps through
with BB Bid (the 2018 help page: "the program waits for you to click
Bid's OK before continuing"). The user takes part only through the bid
box (section 2). Hide Hands only hides cards; it does not stop the engine
bidding for South.

### One call (`MakeBid(seat)`, frm:6061)

1. Clear the disclosure box.
2. Engine inputs (`NewBid`, the `BidBaseDB.Bidder` class; the port's
   `bidbase_engine::lookup::Request`):
   - `DBPathAndName` = `App.Path\BidBase.accdb`.
   - `PlayersHand` = the seat's hand, `AK432-64-QJT3-T2`.
   - `PriorCalls` = `str_AllCalls`, or, when the auction contains "@" (and
     at least two calls were made), the auction from **the first** "@"
     token on: `Mid(auction, pos("@") - 2)`, e.g.
     "1H-P-1S-2D-2S-P-4N@30-P" -> "4N@30-P" (frm:6097-6136).
   - `TrumpSuit` (only in that case) = `Val` of the two characters after
     "@" ("30" -> 30; the engine takes `\ 10` as the trump suit 1-4 and
     `Mod 10` as a second suit, cls:492). A "Q" code (control asking) is
     sent as `n*10+9` with a "Q lookup needs checking." box; an "N" code
     gets "N lookup needs checking.".
   - `OptimumAvg`, `Opt4PardAvg`, `OptimumOrig`, `Opt4PardOrig`,
     `Opt4OppAvg`, `Opt4OppOrig`: double-dummy optimum contracts coded
     `level*10 + strain` (strain 1 = C .. 5 = NT; 34 = 3S, 25 = 2N), for
     this seat, partner, and RHO (section 5). The engine uses them only
     when no rule fits and `OptimumAvg > 0` (`ComputeOptBid`, cls:820,
     1083; not ported). `OptimumAvg` is non-zero only after averaged DDA,
     and the array is never cleared between deals, so stale values from an
     earlier deal can be sent.
   - `Vul`: from `L_Vul` for the bidder's side: NS bidder "Unfav." -> "U",
     "Fav." -> "F", "Both" -> "B"; EW bidder "Unfav." -> "F", "Fav." -> "U",
     "Both" -> "B"; none -> "" (frm:6216-6235). The engine reads "" as "any
     vulnerability", not "neither" (cls:373), so rules that require
     non-vulnerable still match when vulnerable rules would not; with the
     port, `Vul::Unknown` reproduces this and `Vul::Neither` is what was
     meant.
   - `BidToAnalyze` = the analysis request or "".
   - `GetDisclosures` = True. `SaveDealNum` = the deal number.
   - Not set: `PlayersPosition`, `ScoringMethod`, `TestID`, `TestEntry`,
     `CheckAllPcts`, `Activated`, `Combo`, `CancelSearch`. Position comes
     from the opening passes in `PriorCalls`.
3. Read `Bid` (this runs the lookup), `ErrCode`, `Disclosure`, `BON` (when
   a bid came back), `ID`.
4. **No bid**: if `Bid` is "" and `ErrCode` < 4 or = 97 ("no bid found",
   "no matching entry", bad hand length, unknown vulnerability) the call
   is "P" and the error is dropped. Other errors leave an empty call. There
   is no legality check of the engine's call (the port's
   `table::bid_seat` has one).
5. Store the call in the next slot and append it to `str_AllCalls`.
6. If `ErrCode` > 0 and not 2, the disclosure box shows `ErrMsg`.
7. Unless the call is "P", a non-blank `Disclosure` goes in the box.
   Passes never show a disclosure.
8. Redraw the auction grid.

The disclosure text is built by the engine (`GetDiscl`), not by this
program; the port's `Found.disclosure` is the same text (CLI `-d`). Its
first line names the convention, which is what double-click uses. Example
from the 2018 screenshot:

```
1S:  BON:910000.
Strength=Opening.  .  x: S:c H: D: C:
Shape: .  x: S:5+ H: D: C:
Source: 'Stephen Levy web article'
```

Alerts and explanations are not shown separately; whatever the engine
puts in the disclosure is all there is.

### Saving a call to the database (`AddBidToBB`, frm:3562)

Sets `PlayersHand`, `PriorCalls`, `SaveDealNum`, `SaveBid` and then `Job
= "Save"`, which makes the engine insert a row in `Bids` (`SaveBidEntry`,
cls:8416): prior bids without opening passes, `Bid_Order_Number`
"999900" for P, "980000" for D/R, else `9<strain 1-5><level>000`; the
hand's HCP, total points `T<n>`, each suit's length and HCP as exact
values; `Test_Hand` = the hand; `Best_Contract` from `OptimumAvg`;
`Source` = `Deal #<n>`; `DateOfEntry` today; it also writes
`EditPractice.txt` (hand, prior calls). A Rust port reads CSV exports and
has no rule writer; this needs one, or should be left to the Editor.

### End of auction

Checked on the click after the last call: if `str_AllCalls` ends in
"-P-P-P" (three passes after a call, or a fourth initial pass) "Three
consecutive passes." / "Bidding's over." (frm:5958). Nothing else happens:
no final contract, no score, no par, no comparison with the DDA. The
DDA table has been on screen since the deal was shown. The user goes on
with > or Deal.

### PBN auctions

When a PBN event is showing, BidBase still bids every call. The VB meant
to compare each call with the PBN's auction and ask "BB bid <X> and the
PBN player bid <Y>. Refer to the PBN analysis above. / Refer to the
Optimum scores below. Do you want to change to PBN's bid? (Click no to
keep the BidBase bid.)" (frm:6014-6050). It cannot work as written:
`ReDim str_PBNbids(40)` at the first call wipes the parsed auction and
turns the 2-D array into 1-D, and the parser misreads BBA's column layout
(section 7). Reimplement from the intent if wanted.

## 4. Dealing

### The deal number

A deal is a pure function of its number, 1..9,999,999. Dealer and
vulnerability are not part of it. The VB reseeds VB6's `Rnd` with
`Rnd(-dealNumber)` and shuffles with the "old routine" (dealer:479-546;
the 2020 Fisher-Yates at dealer:518-527 is skipped by `GoTo OldRoutine`
because saved entries cite deals made the old way).

**VB6 `Rnd`** is a 24-bit linear congruential generator:

```rust
struct VbRnd { seed: u32 }           // VB starts at 0x50000; irrelevant here

impl VbRnd {
    /// Rnd(x) for x < 0: reseed from the bits of x as a Single, then step.
    fn reseed(&mut self, x: f32) -> f32 {
        let b = x.to_bits();
        self.seed = b.wrapping_add(b >> 24) & 0xFF_FFFF;
        self.next()
    }
    /// Rnd() / Rnd(x) for x > 0.
    fn next(&mut self) -> f32 {
        self.seed = self.seed.wrapping_mul(0x43FD_43FD).wrapping_add(0xC3_9EC3) & 0xFF_FFFF;
        self.seed as f32 / 16_777_216.0
    }
}
```

(`Rnd(-1)` gives 0.2240076, seed 3758214, which is VB's known value.
`Randomize` at frm:6984 does not matter: `Rnd(negative)` replaces the
whole seed.)

**The shuffle** (dealer:490-545):

```rust
fn deal(number: u32) -> [[u8; 13]; 4] {          // [S, W, N, E]
    let mut r = VbRnd { seed: 0x50000 };
    r.reseed(-(number as f32));                    // value discarded
    let mut c: Vec<u8> = (0..=52).collect();       // c[1..=52] = 1..=52
    let passes = 10 + (5.0 * r.next() as f64).floor() as usize; // evaluated once
    for _ in 0..passes {
        for i in 1..=52 {
            let k = (52.0 * r.next() as f64 + 1.0).floor() as usize;
            c.swap(i, k);
        }
    }
    // South = c[1..=13], West = c[14..=26], North = c[27..=39], East = c[40..=52]
    ...
}
```

Compute `52 * Rnd + 1` in double precision (`f64`), not `f32`: in single
precision 10 of the first 20,000 deals come out different, and two deals
cited in the database (20307 and 27415) only reproduce in double.
Checked against the database: 6,461 of the 7,096 rules in `Bids` and
`Garbage` whose `Source` cites "Deal #<n>" have a `Test_Hand` equal to one
of deal n's four hands with this algorithm. The other 635 were not
looked into (hands edited after entry, mixed or rotated deals, or other
generators).

Card numbers are as in section 1 (`suit = (c-1)/13 + 1`, 1 = C .. 4 = S;
`rank = (c-1)%13 + 1`, 1 = A, 2..13 = 2..K). Each seat's 13 cards are then
sorted (`SortHands`, a quicksort, dealer:549) and turned into
`AK432-64-QJT3-T2` strings (`SetUpHands`, dealer:386); the sort does not
change which cards a seat holds.

Test vectors (S / W / N / E):

| Deal | South | West | North | East |
|---|---|---|---|---|
| 1 | J2-T2-AQT9-AQT52 | AQT8-AJ5-J72-876 | 9765-KQ874-5-943 | K43-963-K8643-KJ |
| 24 | AQ9-KT87-T8542-5 | J8742-Q-K96-KJT2 | T53-AJ52-A-AQ643 | K6-9643-QJ73-987 |
| 20307 | QT9642-T9-A9-KJ5 | J8-J86-QJT754-T9 | K5-AK2-K8-AQ7432 | A73-Q7543-632-86 |
| 27214 | KT8-AQJ6-AK932-2 | J9642-5-J8764-86 | Q753-8-Q5-AKJ753 | A-KT97432-T-QT94 |
| 31118 | T32-QT-Q9-KT9643 | A874-AK3-A862-A2 | J965-J4-JT73-J87 | KQ-987652-K54-Q5 |

(24 and 27214 match `Test_Hand`s cited as "Deal #24-N" and "Deal
#27214-N"; 20307 is the double-precision check; 31118 is the number in
`BBPractice.cfg`.)

### Numbering, dealer and vulnerability

- Next deal (>) is number + 1, previous (<) number - 1; the Deal button
  jumps to a typed number. At start the last number is shown again.
- The dealer rotates one seat clockwise with each > unless Keep is
  checked. Deal, Repeat and Dealer leave the auction's dealer as set.
  The dealer letter in references like "Deal #27214-N" is the dealer
  at the time.
- Vulnerability is never cycled; it is whatever the Vulner. button last
  set (or a PBN board's).

### Mix, remix, rotate, deal the rest

- Mix Up Cards / Remix (`ScrambleHands`, dealer:673): pool the cards of
  the chosen seats, then `10 + Int(5*Rnd)` passes of `n` swaps with `k =
  Int(n*Rnd + 1)`, without reseeding, and deal them back in seat order.
  Reproducible only from the generator's state after the deal, not from
  the deal number alone.
- Averaged DDA uses the same scramble (section 5). It honours the seats
  of the last Mix Up Cards if one was made (`int_NumMixedHands` is never
  reset), so after a Mix the averaging scrambles the wrong seats.
- Rotate: section 2.
- Deal the rest: section 2.

## 5. dds.bas and the DDA

### The DLL call

One function, declared as a `Sub` so its return code is ignored
(dds:4):

```
CalcDDtablePBNDDS Lib "dds.dll" Alias "CalcDDtablePBNVB"
    (ByRef ddaDeal As ddTableDealPBN, ByRef res As ddTableResults)
```

- `ddTableDealPBN` = 80 bytes, the PBN deal string, e.g.
  `W:A874.AK3.A862.A2 J965.J4.JT73.J87 KQ.987652.K54.Q5 T32.QT.Q9.KT9643`
  (C `struct ddTableDealPBN { char cards[80]; }`). Built by `ShowDDA`
  (frm:7217): the dealer's letter, ":", then the four hands from the
  dealer clockwise, dashes turned into dots. Over 79 bytes: "String too
  long".
- `ddTableResults` = `int resTable[5][4]`: tricks by strain (0 = S, 1 =
  H, 2 = D, 3 = C, 4 = NT) and declarer (0 = N, 1 = E, 2 = S, 3 = W).
  `CalcDDTablePBN` (dds:163) transposes the array to undo VB's
  column-major layout, so afterwards `res.resTable(strain, hand)` means
  the same as in C. `ChDir App.Path` first so the DLL is found.
- The alias `CalcDDtablePBNVB` is not a standard DDS export; the
  bundled `dds.dll` (Sept 2019) is a build with it. The comment cites
  DDS 2.9.0's DLL description. Every other type in `dds.bas` (`boards`,
  `parResults`, `DDSInfo`, ...) and the `IntTo*` helpers are unused.
- A Rust port can use any DDS binding for a full 20-result table (the
  same thing rusty-bidding-bot's `rbb dd` computes).

### Where results go

`ShowDDA(DDA_Only)` (frm:7161) runs on every new deal, Mix, PBN board and
user deal (not on Repeat):

1. One DDS table for the deal.
2. For each declarer, the five trick counts are kept as
   `int_OptByStrain(dds_hand + 1, strain, 0)` (N = 1, E = 2, S = 3, W = 4;
   strain 1 = C .. 5 = NT) and formatted into `str_DDAOutput`.
3. Unless `DDA_Only`: HCP, LoTT, LTC and the five lines of `cmb_DDA`
   (section 1); `DDAresults.txt`; `CalcOpt(0)`; averaged DDA if that
   option is checked.

`CalcOpt` (frm:7937) picks each declarer's "optimum": each strain's trick
count is scored with a fixed table that ignores vulnerability (fewer than
7 tricks 0; 7-8 tricks in C/D `50 + (t-6)*20`, H/S `50 + (t-6)*30`, NT
`60 + (t-6)*30`; 9 NT 400; 10: 130 / 420 / 430; 11: 400 / 450 / 460; 12:
920 / 980 / 990; 13: 1440 / 1510 / 1520 for minors / majors / NT). The
best score wins; for indices 3 and 4 a tie between C and D or H and S
goes to the longer combined fit (the higher suit if equal), for 1 and 2
the higher strain wins ties. The optimum is `(tricks - 6) * 10 + strain`,
0 below one level. `CalcOpt(0)` fills `int_OptOrig`.

### Averaged DDA (`AverageDDA(n)`, frm:7561): medium

Options "11-Deal Avg. DDA" (n = 11) and "101-Deal Average" (n = 110).
Shows `t_Calculating` while it runs.

1. Keep North-South, scramble East-West's cards n times (DDS each time);
   then restore and keep East-West, scramble North-South n times.
2. For the kept pair, per declarer and strain, the median trick count:
   the highest t such that t or more tricks came up in more than n/2
   layouts (threshold 5.5 or 55).
3. Lines added to `cmb_DDA` (index 6 on): `Median DDA: 5.5` (the
   threshold is printed), then `N:`, `S:`, `E:`, `W:` each followed by
   five 3-wide numbers for C D H S N. The 110-run puts `Median DDA (11)
   C  D  H  S  N  (101) 55` on line 6 and appends its numbers after the
   11-run's.
4. The optimum per seat from the medians is `int_OptAvg`, sent to the
   engine as `OptimumAvg` etc.
5. The medians are appended to `DDAresults.txt`.

Seat numbering is mixed up here: `ShowDDA` stores DDS declarers (N, E,
S, W = 0..3) in `int_DDA` and `int_OptByStrain`, while `AverageDDA`,
`CalcOpt`'s fit test and `MakeBid` index them by program seat (S, W, N,
E = 1..4). The result is that the median rows show North's numbers
against S and South's against N, East's against W and West's against E,
and the engine's optimum for a seat comes from the seat opposite (same
partnership). A port should index everything by one seat type.

### How DDA feeds the UI

Only as the table under the hands and, through the optimums, as the
engine's last-resort `ComputeOptBid`. It is not compared with the
auction, and there is no par score.

## 6. BidBase Eval.bas

3,095 lines, module `BBEval`, no `Option Explicit`. It is a copy of the
engine's hand analysis (`AnalyzeHand`, `Calculate_Winners`,
`Calculate_LosingTrickCount`, `Calculate_QuickTricks`,
`Calculate_Stoppers`, `Calculate_Shape`, `GetOtherPlayersData`) plus the
hand-valuation evaluator `Eval_A_Hand(sysName)` (eval:176), which scores a
hand against the rows of a `HandVal` table (BB-HEP style: value, flag,
position, hands, fit, suit, cards and length specs per row) shown in a
`DBGrid1` bound to `Data1`. Its only public entry points are
`ComputeBBVal` (eval:150, ends in a "Stop ComputeBBVal 1" box) and
`Eval_A_Hand`. Neither is called by the form or the dealer module, and
`DBGrid1`/`Data1` do not exist on `PlayScr`. **Dead code**; the program
uses only its constants `Ace = 1`, `King = 13` (and `Queen`, `Jack`, ...)
(eval:7-18). Do not port it. The 2026 engine dropped valuation systems
too (`ENGINE-2026.md`).

## 7. Files

All in the program folder unless noted; ANSI (Windows-1252) text, CRLF.
VB `Print #` writes a number with a leading space for the sign and a
trailing space (" 31118 "); `Print #` with commas pads to 14-character
zones; `Write #` writes `#TRUE#`/`#FALSE#` and quoted strings.

### BBPractice.cfg

Written on exit (`Form_Unload`, frm:3401); read at start (`LoadCfg`,
frm:6931; if missing, `BBPractice new.cfg`, else "...BBPractice.cfg is
missing. Default settings will be used."). The current file:

```
 31118 
 17160         4830 
#TRUE#
#TRUE#
 2 
 0 
""
 0 
```

| Line | Written | Read as |
|---|---|---|
| 1 | current deal number | deal number - 1 (the start-up deal adds 1) |
| 2 | window Left, Top in twips | position |
| 3 | `Not AutoBid` | `bln_AutoBid` |
| 4 | `Not HideHands` | `bln_HideHands`, then `mHide_Click` flips it back |
| 5 | dealer 1-4 (S W N E), 0 if none | dealer (0 -> South) |
| 6 | Keep checkbox 0/1 | Keep |
| 7 | PBN event file (always "") | not read |
| 8 | PBN board number | not read |

AutoBid is stored inverted because start-up was meant to run
`mAutoBid_Click`, but that call is commented out (frm:6994), so each
session starts with Auto-Bid flipped while the menu still shows it
checked. A port should store plain values (TOML) and may read this
format once to migrate.

The 2018 `BBPractice new.cfg` in `sourceforge-2018/` has an older layout
(deal number, a 1, position, `False`, `True`, 0, 0, and a "Start with"
string) and would not load correctly with this code.

### Last Deal Number.txt

One line, `Print #` of the deal number (" 31118 "). Written by > / < and
by leaving PBN/user mode; read at start before `BBPractice.cfg`, which
wins when present.

### DealerPosition.txt

Two lines, the window's Left and Top in twips, written on every BB Bid
to `App.Path & "DealerPosition.txt"` (no backslash, so it lands next to
the program folder as `<folder>DealerPosition.txt`). The engine reads the
same path in `ComputeOptBid` to place its "Click Ok to save computed
call..." InputBox (cls:10910). A port needs it only if it ports
`ComputeOptBid`.

### EditPractice.txt

Written by Edit Bid (section 2) and by the engine's `SaveBidEntry`; read
by the BidBase Editor's import. Three lines:

```
KT8-AQJ6-AK932-2
1C-2H
Deal #27214-N ID=1234 BON=910000
```

### DDAresults.txt

Written by `ShowDDA` in the current directory: four lines, the
14-character DDA strings for S, E, N, W (in that order). `AverageDDA`
appends four lines of medians for S, W, N, E. The engine's `GetDDA`
(cls:8680) reads it back (12 values, with its own seat order) for
`ComputeOptBid`. The 2018 code wrote the optimum array here for the same
purpose (frm:6185-6211, now `If 1 = 2`).

### Notes\SavedDeals.txt (Save Deal)

Appended (deal 31118 with an invented auction):

```


Deal #31118
N:  J965-J4-JT73-J87
S:  T32-QT-Q9-KT9643
W:  A874-AK3-A862-A2
E:  KQ-987652-K54-Q5
Dealer: W
Bidding:
P-1C-P-1H
P-1N-P-P
P
<the cmb_DDA lines>

```

Calls joined with "-", a new line after every fourth call and after the
last.

### UserEnteredHands\<name>.txt (user deals)

52 lines, one card number 1-52 each (`Print #`, " 40 "), 13 per seat in
the order South, West, North, East. Within a seat the order is whatever
the program had; a reader should treat each block as a set. `LoadDeal`
(frm:4274) builds the hand strings assuming descending card numbers and
adds only one dash across a void; the caller then rebuilds them with
`SetUpHands`, which is what a port should do.

### PBN events (File, Load PBN Event; `mLoadPbn_Click` frm:4339, `GetPbnDeal` frm:4455)

- File dialog in `PBNs\` (`*.pbn`); "The deal to be loaded will / replace
  the current deal. / Continue?".
- Index pass: every line is numbered; a blank line marks a board start.
  Each `[Board` adds a list entry `Line <n>: Round-Board= <round><board>
  <event>` (n = the preceding blank line). Errors: a line over 1000
  characters ("This PBN file cannot be loaded. It does not have a Carriage
  Return and Line Feed between each board.") or `[Board` with no blank
  line before it after line 9 ("...does not have a blank line between
  each board.").
- Board pass: skip to the line, read until the next blank line once a
  dealer has been seen. Tags are found with `InStr` (anywhere on the line)
  and read at fixed columns, so text after the tag (as in `Gerber.pbn`)
  does no harm:
  - `[Dealer "S"]`: column 10.
  - `[Vulnerable "NS"]`: columns 14-15: "No" -> "", "Al" -> "Both",
    "NS" -> "Unfav.", "EW" -> "Fav.", anything else is copied ("Bo" from
    "Both", which then matches no code and is sent as no
    vulnerability). "Love" and "-" are not handled.
  - `[Deal "N:..."]`: first seat at column 8, hands from column 10, four
    16-character hands separated by spaces, clockwise from the first seat.
    Assumes exactly 13 cards per hand.
  - `[Board`, `[Round`, `[Event`: board number, round label, event name.
  - `[West`, `[North`, `[East`, `[South` longer than 12 characters: listed
    in the disclosure box under "      Players: ".
  - `[Score` (quotes stripped) and `[Result` (whole line, overwrites):
    kept for the PBN comparison prompt.
  - `[Auction`: following lines until one containing "["; see below.
  - A line `{` starts commentary up to a line `}`; ". " and "? " become
    line breaks; shown by Bid w/ Analysis.
- `L_Deal` becomes `<round><board> - <file name without .pbn>`, on two
  lines if wider than 280 pixels. Arrows step through the list; File,
  Load PBN Deal # picks from it. The deal is shown and the DDA computed;
  the auction is not replayed.
- The auction parser (frm:4588-4604) replaces "Pass" with " P", "NT"
  with "N", "XX" with " R", "X" with " D", shrinks runs of spaces before
  "P", then reads calls at columns 1, 4, 7 and 10 (two characters each).
  It works for single-space lines like "Pass 1NT X XX" and fails on BBA's
  six-column layout: "Pass  Pass  Pass  1C" gives P, P, P, "1";
  "1C    Pass  1H    Pass" gives 1C, P, "1", "". A port should tokenise
  on whitespace and drop `=n=` notes and `$n` NAGs.

`Gerber.pbn` (295 boards) is a sample event: Bridge Bidding Analyser
output (SAYC both sides, 2023-10-16) with Nelson's notes typed to the
right of tag and auction lines, e.g. `[Auction "S"]   BidBase opens 1D
insead of 1C.` It has no `{}` commentary.

### Convention practice files (unfinished feature)

- `PracticeDeals\<name>` (no extension): auction specs, one per line
  (section 2).
- `<name>.Deals`: random-access file of 6-byte little-endian records.
  Record 1: `NumEntries` (i32), `LastNumUsed` (i16). Records 2..: deal
  number (i32), dealer (i16, 1 = S, 2 = W, 3 = N, 4 = E). Deleting copies
  the last record over the deleted one and decrements `NumEntries`. The
  code checks for it in `PracticeDeals\Deals\` (frm:4100, 2849) but opens
  it by bare name, i.e. in the current directory (frm:3125, 3169).
- `<name>.Txt`: appended for each found deal, a `Write #` line `entry,
  deal number, dealer, "dealer's hand"`.
- `Notes\Practice_<name>.txt`: HTML fragments to paste into a quiz page
  (`MakePracticeWebPage`, frm:8103). Every space in the whole path is
  turned into "_" (frm:8138), which breaks paths like `C:\Program Files
  (x86)\...`. Per item (values illustrative):

  ```html
  <BR>7. &nbsp;Deal #3561
  <UL>Your hand: <B>AK2-...</B>
  <BR>Prior bids: &nbsp;<B>(1C)-D-(P)</B>
  <SPAN ID="prompt7" STYLE=''" COLOR=BLACK>
      <SPAN OnClick="prompt7.style.display='none'">
          <SPAN OnClick="hidden7.style.display=''">
              <SPAN id="GreenHand"><BR>
                   <U><B>Your bid</B></U>
             </SPAN>
         </SPAN>
     </SPAN>
  </SPAN>
  <SPAN ID="hidden7" STYLE="display:none" COLOR=BLACK>
  <BR>BidBase: &nbsp;<SPAN style="color:#FF0000"><U><B>2C</B></U><SPAN style="color:#000000"> &nbsp;(DDA:   C=9 &nbsp;D=10 ...)</SPAN></SPAN>
  </SPAN>
  </UL>
  ```

  The opponents' calls in "Prior bids" are in parentheses (the first and
  last of three). The 2nd-bid variant adds "Your 1st bid", "Next 3 bids",
  a hidden "Your 2nd bid", "Pard's hand" and "DDA". A "Stop
  MakePracticeWebPage 1" box pops up on every item.

### Other paths the program uses

`Notes\*.htm` (help, convention notes), `BidBase File Viewer.exe`,
`BidBase Editor.exe`, `c:\windows\notepad.exe`, `Qcard32.dll`, `dds.dll`,
`BidBase.accdb` (through the engine).

## 8. Porting order

### MVP

1. **Deals.** `VbRnd` + the old shuffle (section 4), with the test
   vectors as unit tests; hands as `AK432-64-QJT3-T2` and PBN.
2. **Table state.** Deal number, dealer, vulnerability (none / NS / EW /
   both), Keep; > and < (dealer ± 1 seat), Deal (jump to a number),
   Repeat, Dealer. Persist the deal number, dealer, Keep, hide-hands and
   window position in a config file.
3. **Hands.** Draw the four hands as fans (N/S across, W/E down, 21/27
   pixel steps for 71 x 96 cards, or a text layout); Hide Hands for W, N,
   E. Card images: any free set; Qcard32 is Windows-only.
4. **Engine.** For each call: `Request { hand, prior_calls: tokens
   joined with "-", opening passes included, vul for the bidder's side,
   trump }`, cutting the auction at **the first** "@" token as the VB does
   (frm:6101). The port's `table::key_for` cuts at the last one
   (`rposition`, `crates/engine/src/table.rs:48`); 22 keys in the 2026
   rules hold two or more "@" tokens and are keyed from the first
   ("5D@CC-P-5H@CS-P-6S-P"), so `key_for` should use `position`. No
   rule: Pass. Keep the full token ("4N@30") in the auction, show "4N".
   Check legality.
5. **Auction grid** S W N E, dealer's column first, P/D/R as letters.
   **BB Bid** makes the next call; three passes after a call (or four)
   ends it.
6. **Disclosure** panel with the engine's text (none for Pass) and error
   text.
7. **Bid box.** For a first version: clicking a call when it is a seat's
   turn means "my call for this seat"; show BidBase's call and whether
   they agree, and let the user keep either (keeping the auction string
   in step with the grid, unlike the VB).
8. **DDA table.** One DDS table per deal and the five-line block of
   section 1 (DDA, HCP, LoTT, LTC). Optionally par, which the VB never
   had.

### Later

- Edit Bid / `EditPractice.txt` export (cheap; useful only with the
  Editor), and Bid w/ Analysis as a view of `Outcome.tried`.
- Enter Hands with a card picker and text boxes; Deal the rest done
  properly; save/load user deals (read the 52-number format).
- Mix Up Cards / Remix (any RNG will do; VB-exactness is not needed), and
  Rotate done on one seat array.
- PBN events: board list, dealer/vulnerability/players/commentary, a
  whitespace auction parser; optionally replay and compare the PBN
  auction.
- Save Deal to a text file.
- Averaged DDA (11 / 110 layouts per pair, medians, optimums), with one
  consistent seat index. Only needed for the engine's `ComputeOptBid`,
  which is not ported.
- Adding rules from the bid box (`Job = "Save"`): needs a rule writer for
  the CSV data or the Editor.
- Convention practice (sections 2 and 7): the largest feature and
  unfinished in the VB; design it fresh on top of the deal generator
  (auction-spec matcher with "A", "x<call>" and "|", search over deal
  numbers and dealers, saved (deal, dealer) lists, a quiz view, HTML
  export).
- Help, Run Editor, Stay on Top, File Viewer links: platform glue.

## 9. The Rust port (`crates/practice`)

Started 2026-10-13. The screen is Nelson's form, read from the `.frm`
and `.frx` at build time (`crates/forms`) and drawn with egui
(`crates/vbui`: his controls, colours, fonts, z-order, menus). The
handlers in `crates/practice/src/lib.rs` follow his, named in their doc
comments. His message boxes and prompts are his own frames (`f_Msg`,
`f_Inbox`) sized as `MyMsgBox` sizes them. The pieces:

| Piece | His | Port |
|---|---|---|
| Screen | `PlayScr` form | `bidbase_vbui::Screen` over the same `.frm` |
| Deals | `Rnd(-n)` + old shuffle | `dealer.rs`, same deals (test vectors above) |
| Cards | Qcard32.dll | `cards.rs`, drawn, same 71 x 96 size and fan steps |
| Calls | `BidBaseDB.Bidder` | `bidbase_engine::table::next_call_with`, 2026 rules compiled in |
| DDA | dds.dll | bridge-solver (`solve_dd_table`), off the drawing thread on native |
| Settings | `BBPractice.cfg` | eframe storage (a file natively, local storage on the web) |

Run it with `./dev-build.sh run --release -p bidbase-practice`, or in the
browser at bridge-craftwork.com/bidbase/practice/ (`web/build.sh`).
`crates/practice/tests/screens.rs` draws its screens offscreen; with
`BIDBASE_SHOTS=<dir>` it writes them as PNGs.

**Ported:** start-up (the saved deal, its DDA, the dealer's first call);
Deal (prompt for a number), < and >, Repeat, Vulner., Dealer, Keep; BB
Bid; the bid box (both cases); Edit Bid (shows what `EditPractice.txt`
would hold); Bid w/ Analysis (the rules tried for the watched call, in the
disclosure box); the auction grid; the disclosure; the DDA table; Save
Deal (appends to `~/BidBase/SavedDeals.txt` natively, shown in the
disclosure box); Auto-Bid, Hide Hands, Stay on Top, Mix Up Cards, Remix,
Rotate; the help and conventions pages (Nelson's site); Enter and Alt
keys.

**Departures**, each a fix of something in section 2-5 or a gap:

- No vulnerability is sent as neither side vulnerable, not as blank
  (which the engine reads as "any").
- After "No" in the bid box, the auction sent to the engine has the
  user's call, as the grid shows (his kept BidBase's). Case 1's "Yes"
  replaces BidBase's last call (his removed every copy of it from the
  auction string) and does not add a rule: the port has no rule writer, so
  its message says that adding the call to the rules needs the Editor. Its
  question reads "Do you want your bid of X to replace BB's bid of Y?".
- Illegal calls are refused (BidBase's become Pass, with a note).
- < steps the dealer back one seat; Rotate turns one seat array and
  redoes the DDA; the E-S rotation uses three seats, not 23.
- The DDA's self-sufficient-suit test uses honours for both sides.
- When the auction ends, the disclosure box starts with the final
  contract (his showed nothing).
- Auctions longer than six rounds scroll in the grid (his stopped).
- The deal prompt takes a number only (his "user" did nothing).
- Card backs for hidden hands are drawn red, as Qcard32's design 3.

**Not yet:** Enter Hands (card picker, text entry, Deal the rest,
user-deal files), Load User Deal, PBN events, the averaged DDA, Run
Editor, and convention practice (unfinished in his source too). Their
menu items are greyed out or say "not in this port yet".
