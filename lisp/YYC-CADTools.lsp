;;; ===========================================================================
;;; YYC-CADTools.lsp  -  YYC CAD Export toolkit (ActiveX edition)
;;; DIALOG Design Technology  -  companion to the YYC CAD Export guide
;;; https://shadiaalkhaledi-dot.github.io/CADexport/
;;;
;;; Built on ActiveX (the vla-/vlax- functions): every tool works on a document
;;; OBJECT, so the same code runs on the drawing you have open or on drawings
;;; YYCBATCH opens through ActiveX - no command-line macros, no script files.
;;; Needs full AutoCAD on Windows (ActiveX is not available in LT or on Mac).
;;;
;;; Load:  APPLOAD -> this file (add it to the Startup Suite to keep it loaded)
;;; First: YYCSETUP - points the tools at YOUR copy of the YYC kit files.
;;;        Settings are stored per Windows user; nothing is hard-coded.
;;;
;;; Commands, in process order
;;;   Step 2   YYCRENAME      Rename exported DWGs in a folder to <DwgNo><Sheet>.dwg
;;;   Step 3   YYCPAGESETUP   Copy "YYC - Titleblock" page setup from the .dwt to all layouts
;;;   Step 4   YYCSCHEDULES   Pick a viewport: its contents go to paper space, then delete it
;;;   Step 5   YYCGRIDIN      Insert the YYC grid (04grid-dtb.dwg) as a block at 0,0,0
;;;            YYCALIGNREC    Do the MOVE/ALIGN once with 4 clicks; saves it to the profile
;;;            YYCALIGNAPPLY  Apply the saved move to this drawing (or Align in YYCBATCH)
;;;            YYCVPUCS       Viewport UCS to match its view (fix for drawings done earlier)
;;;            (YYCALIGNUSE   hidden - borrow another profile's move; not in the instructions)
;;;   Step 6   YYCGRIDOUT     Delete the grid block and its definition
;;;            YYCCLEAN       Delete empty text, PurgeAll x3, AuditInfo (fix)
;;;            YYCFIND0       Report named objects containing "$0$" (bind leftovers)
;;;   Step 7   YYCVPLAYERS    Put every viewport on its own VIEWPORT# no-plot layer
;;;            YYCVPLOCK      Lock every viewport and list its scale
;;;   Step 8   YYCLWDEFAULT   Set every layer's lineweight to Default
;;;            YYCLAYSYNC     Layer colour/linetype/lineweight to match the YYC Layer Reference
;;;            YYCLAYXL       Send layers to the Layer Mapper workbook (Excel) to choose YYC layers
;;;            YYCLAYMAP      Apply the workbook choices (rename or merge), remember them
;;;            YYCLAYWALK     Walk what is left one layer at a time: show it, pick a YYC layer, map
;;;   Setup    YYCMAKEREF     Build the YYC Layer Reference DWG + DWS from the standard CSV (once per kit)
;;;            YYCLAYEXPORT   Write layers to YYC_LayerExport.csv, checked against the
;;;                           YYC Layer Reference
;;;            YYCZERO        Report and select objects on 0 / DEFPOINTS (report only)
;;;   Step 9   YYCQA          CADD Manual 5.15 check - report only, changes nothing
;;;            YYCFINAL       Lock viewports, main layout active, zoom extents
;;;   Step 10  YYCFILELIST    Write filelist.txt for a folder and check file names
;;;   Any      YYCBATCH       Run a chosen sequence of the above over a folder
;;;            YYCSETUP (first time) / YYCPROFILES (everything else about profiles)
;;;   Help     YYCHELP        Every command in process order, with a Run button
;;;            A YYC.cuix ribbon tab next to this file is loaded automatically.
;;;
;;; Profiles: one set of kit files per project or area (grid, template, layer
;;; reference...). YYCPROFILES switches; YYCGRIDIN lets you pick the grid each time.
;;; ===========================================================================

(vl-load-com)
(setq *yyc-version* "1.1")
(setq *yyc-modifiers* '("DEMO" "EXST" "FUTR" "MOVE" "NEWW" "NICN" "NPLT" "PRPS" "RELO" "TEMP"
                        "ABDN" "RMVD" "PATT" "SYMB" "TEXT" "IDEN" "EQPM" "ELEV"))

;;; ---------------------------------------------------------------------------
;;; Core helpers
;;; ---------------------------------------------------------------------------

(defun yyc:acad () (vlax-get-acad-object))
(defun yyc:doc () (vla-get-ActiveDocument (yyc:acad)))

;; Settings live in named PROFILES (e.g. "24C024 DTB", "ITB"), one set of kit
;; files per project or area. Profile / Profiles / ToolsPath are shared.
(setq *yyc-keys* '("Description" "Scales" "SchedMargin" "StdCSV" "LayerBook" "Kit" "Template" "PageSetup" "Grid" "GridBlock" "LayerRef" "LinFile" "LayerMap" "QABook" "FDTemplate" "DwgNo" "CTB" "Fonts"))

(defun yyc:profile ( / p) (setq p (getenv "YYC_Profile")) (if (and p (/= p "")) p "Default"))
(defun yyc:pkey (key profile)
  (if (member key '("Profile" "Profiles" "ToolsPath"))
    (strcat "YYC_" key)
    (strcat "YYC_P_" (vl-string-translate " " "_" profile) "_" key))
)
(defun yyc:get-in (profile key default / v)
  (setq v (getenv (yyc:pkey key profile)))
  (if (and v (/= v "")) v default)
)
(defun yyc:set-in (profile key val) (setenv (yyc:pkey key profile) (if val val "")))
(defun yyc:get (key default) (yyc:get-in (yyc:profile) key default))
(defun yyc:set (key val) (yyc:set-in (yyc:profile) key val))
(defun yyc:profiles ( / v)
  (setq v (getenv "YYC_Profiles"))
  (if (and v (/= v "")) (yyc:split v "|") (list (yyc:profile)))
)
(defun yyc:add-profile (name / ps)
  (setq ps (yyc:profiles))
  (if (not (member (strcase name) (mapcar 'strcase ps))) (setenv "YYC_Profiles" (yyc:join (append ps (list name)) "|")))
)

(defun yyc:batch-p () (= *yyc-batch* T))

;; in a batch every message is prefixed with the drawing it belongs to
(defun yyc:msg (s) (princ (strcat "\n" (if *yyc-cur* (strcat "[" *yyc-cur* "] ") "") s)))

(defun yyc:trim (s) (vl-string-trim " \t\r\n" s))

(defun yyc:split (str delim / pos res)
  (while (setq pos (vl-string-search delim str))
    (setq res (cons (substr str 1 pos) res))
    (setq str (substr str (+ pos (strlen delim) 1)))
  )
  (reverse (cons str res))
)

(defun yyc:join (lst delim / out)
  (setq out "")
  (foreach x lst (setq out (if (= out "") x (strcat out delim x))))
  out
)

(defun yyc:inc (key alist / p)
  (if (setq p (assoc key alist))
    (subst (cons key (1+ (cdr p))) p alist)
    (cons (cons key 1) alist)
  )
)

(defun yyc:err-p (x) (vl-catch-all-error-p x))
(defun yyc:try (fn args / r) (setq r (vl-catch-all-apply fn args)) (if (yyc:err-p r) nil r))

(defun yyc:ask (msg default / s)
  (setq s (getstring T (strcat "\n" msg (if (and default (/= default "")) (strcat " <" default ">") "") ": ")))
  (if (= s "") default s)
)

(defun yyc:yes (msg default / k)
  (initget "Yes No")
  (setq k (getkword (strcat "\n" msg " [Yes/No] <" default ">: ")))
  (= (if k k default) "Yes")
)

(defun yyc:stamp () (menucmd "M=$(edtime,$(getvar,DATE),YYYY-MO-DD HH:MM)"))
(defun yyc:stamp-file () (menucmd "M=$(edtime,$(getvar,DATE),YYYYMODD-HHMMSS)"))

;; --- document facts, read from the document object (works for any open doc)
(defun yyc:dname (doc) (vla-get-Name doc))
(defun yyc:dbase (doc) (vl-filename-base (vla-get-Name doc)))
(defun yyc:dfolder (doc / p) (setq p (vla-get-Path doc)) (if (= p "") "" (strcat p "\\")))
(defun yyc:dfull (doc) (vla-get-FullName doc))
(defun yyc:active-p (doc) (equal (strcase (yyc:dfull doc)) (strcase (yyc:dfull (yyc:doc)))))

;; Pick a folder with the normal Windows file dialog (you can paste a path into
;; File name). Open the folder and pick ANY file in it - the folder is what's kept.
;; start: folder to open in (optional).
(defun yyc:browse-folder (msg / start f)
  (setq start (cond ((and (setq start (yyc:get "Kit" nil)) (wcmatch (strcase msg) "*KIT*") (vl-file-directory-p start)) start)
                    ((/= (getvar "DWGPREFIX") "") (vl-string-right-trim "\\" (getvar "DWGPREFIX")))
                    (T nil)))
  (yyc:msg (strcat msg ": go into the folder (or paste its path in File name), pick any file in it, Open."))
  (setq f (getfiled (strcat msg " - pick any file in the folder") (if start (strcat start "\\") "") "" 0))
  (if f (vl-filename-directory f))
)

(defun yyc:dwgs-in (folder / files)
  (setq files (vl-directory-files folder "*.dwg" 1))
  (if files (vl-sort files '(lambda (a b) (< (strcase a) (strcase b)))) nil)
)

;; every kind of kit file the tools use: key, label, search patterns, extension
(setq *yyc-file-kinds*
  '(("Template" "Titleblock template (.dwt)" ("*.dwt") "dwt")
    ("Grid"     "Grid drawing (04 DTB domestic, 20 ITB international ...)" ("*grid*.dwg") "dwg")
    ("StdCSV"   "YYC standard layer list (YYC_Standard_Layers.csv, from the CADD Manual)" ("*Standard*Layer*.csv") "csv")
    ("LayerBook" "YYC Layer Mapper workbook (.xlsx)" ("*Layer*Mapper*.xlsx") "xlsx")
    ("QABook"   "YYC QA Report template (.xlsx)" ("*QA*Report*.xlsx") "xlsx")
    ("FDTemplate" "File Description template (.docx)" ("*File Description*.docx") "docx")
    ("LayerRef" "YYC Layer Reference drawing (used if there's no standard CSV)" ("*layer*.dwg" "*.dws") "dwg")
    ("LinFile"  "CAA linetype file" ("*.lin") "lin")
    ("LayerMap" "Layer map CSV (old,new) - optional, YYCLAYMAPMAKE writes one" ("*LayerMap*.csv" "*Layer*Map*.csv") "csv")))

;; every file under folder matching any pattern (4 levels down), sorted by name
(defun yyc:find-all (folder patterns depth / res)
  (if (and folder (vl-file-directory-p folder))
    (progn
      (foreach pat patterns
        (foreach f (vl-directory-files folder pat 1)
          (if (not (member (strcase (strcat folder "\\" f)) (mapcar 'strcase res)))
            (setq res (cons (strcat folder "\\" f) res))))
      )
      (if (> depth 0)
        (foreach sub (vl-directory-files folder "*" -1)
          (if (not (member sub '("." "..")))
            (foreach f (yyc:find-all (strcat folder "\\" sub) patterns (1- depth))
              (if (not (member (strcase f) (mapcar 'strcase res))) (setq res (cons f res)))))))
    )
  )
  (vl-sort res '(lambda (a b) (< (strcase (vl-filename-base a)) (strcase (vl-filename-base b)))))
)

;; Pick one file of a kind: numbered list of everything found in the kit folder,
;; B to browse anywhere, Enter keeps the current one. Saves the choice.
(defun yyc:choose-file (key / kind cands cur i ans pick)
  (setq kind (assoc key *yyc-file-kinds*) cur (yyc:get key nil))
  (setq cands (yyc:find-all (yyc:get "Kit" nil) (nth 2 kind) 4))
  (if (and cur (findfile cur) (not (member (strcase cur) (mapcar 'strcase cands)))) (setq cands (cons cur cands)))
  (if (not cands)
    (progn
      (yyc:msg (strcat (nth 1 kind) ": none found in the kit folder."))
      (if (yyc:yes "  Browse for one?" (if (member key '("LayerMap" "LinFile")) "No" "Yes"))
        (setq pick (getfiled (nth 1 kind) (yyc:get "Kit" "") (nth 3 kind) 0))))
    (progn
      (yyc:msg (strcat (nth 1 kind) " - profile \"" (yyc:profile) "\":"))
      (setq i 0)
      (foreach c cands
        (setq i (1+ i))
        (yyc:msg (strcat (if (and cur (= (strcase c) (strcase cur))) "  * " "    ") (itoa i) ". "
                         (vl-filename-base c) (vl-filename-extension c) "   (" (vl-filename-directory c) ")")))
      (setq ans (getstring (strcat "\nNumber, B to browse, Enter = " (vl-filename-base (if cur cur (car cands))) ": ")))
      (cond
        ((= ans "") (setq pick (if cur cur (car cands))))
        ((= (strcase ans) "B") (setq pick (getfiled (nth 1 kind) (yyc:get "Kit" "") (nth 3 kind) 0)))
        ((and (> (atoi ans) 0) (<= (atoi ans) (length cands))) (setq pick (nth (1- (atoi ans)) cands)))
        (T (yyc:msg "Not a choice - kept the current one.") (setq pick cur))
      )
    )
  )
  (if pick
    (progn (yyc:set key pick) (yyc:msg (strcat "  Saved " key ": " pick)))
    (yyc:msg (strcat "  " key ": not set (skipped).")))
  pick
)

;; stored file path; asks when missing - but never in the middle of a batch
(defun yyc:need-file (key title ext / p)
  (setq p (yyc:get key nil))
  (cond
    ((and p (findfile p)) p)
    ((yyc:batch-p) (yyc:msg (strcat "SKIPPED: " key " is not set or the file is missing. Run YYCSETUP.")) nil)
    ((assoc key *yyc-file-kinds*) (yyc:choose-file key))
    (T (setq p (getfiled title (yyc:get "Kit" "") ext 0)) (if p (yyc:set key p)) p)
  )
)

(defun yyc:layer-ci (layers name / hit)
  (vlax-for L layers (if (= (strcase (vla-get-Name L)) (strcase name)) (setq hit L)))
  hit
)

(defun yyc:paper-layouts (doc / res)
  (vlax-for lay (vla-get-Layouts doc)
    (if (/= (strcase (vla-get-Name lay)) "MODEL") (setq res (cons lay res)))
  )
  (vl-sort res '(lambda (a b) (< (vla-get-TabOrder a) (vla-get-TabOrder b))))
)

;; blocks that belong to this drawing (not xrefs, not xref-dependent)
(defun yyc:own-block-p (blk)
  (and (= (vla-get-IsXRef blk) :vlax-false) (not (vl-string-search "|" (vla-get-Name blk))))
)

;; run a tool on the active drawing, wrapped in one undo step
(defun yyc:run (fn / doc r)
  (setq doc (yyc:doc) *yyc-cur* nil)
  (yyc:msg (strcat "Profile: " (yyc:profile)))
  (vla-StartUndoMark doc)
  (setq r (vl-catch-all-apply fn (list doc)))
  (vla-EndUndoMark doc)
  (if (yyc:err-p r) (yyc:msg (strcat "ERROR: " (vl-catch-all-error-message r))))
  (princ)
)

;; read another drawing without opening it (ObjectDBX, part of ActiveX)
(defun yyc:dbx-open (path / dbx r)
  (setq dbx (vl-catch-all-apply 'vla-GetInterfaceObject
              (list (yyc:acad) (strcat "ObjectDBX.AxDbDocument." (itoa (atoi (getvar "ACADVER")))))))
  (cond
    ((yyc:err-p dbx) (yyc:msg "ObjectDBX is not available on this machine.") nil)
    ((yyc:err-p (setq r (vl-catch-all-apply 'vla-Open (list dbx path))))
     (yyc:msg (strcat "Could not read " path ": " (vl-catch-all-error-message r)))
     (vlax-release-object dbx) nil)
    (T dbx)
  )
)
(defun yyc:dbx-close (dbx) (if dbx (vlax-release-object dbx)))

;;; ---------------------------------------------------------------------------
;;; Setup and profiles
;;; ---------------------------------------------------------------------------

(defun yyc:switch-profile (name copy / old)
  (setq old (yyc:profile))
  (yyc:add-profile old)
  (yyc:add-profile name)
  (if copy (foreach k *yyc-keys* (if (not (yyc:get-in name k nil)) (yyc:set-in name k (yyc:get-in old k nil)))))
  (setenv "YYC_Profile" name)
)

;; YYCSETUP opens the same window - first time: set the Kit folder (...), click
;; "Fill empty files from the kit folder", type a Description, Save + use.
(defun c:YYCSETUP ()
  (yyc:msg "Setup happens in the profiles window: Kit folder (...) > Fill empty files from the kit folder > Description > Save + use.")
  (c:YYCPROFILES)
)

;; the old command-line walkthrough, kept for anyone who prefers it
(defun c:YYCSETUPCMD ( / name kit)
  (yyc:msg (strcat "YYC CAD Tools setup. Profiles: " (yyc:join (yyc:profiles) ", ")))
  (yyc:msg "A profile is one set of kit files - make one per project or area (e.g. \"24C024 DTB\", \"ITB\").")
  (setq name (yyc:ask "Profile to set up" (yyc:profile)))
  (if (/= (strcase name) (strcase (yyc:profile)))
    (yyc:switch-profile name (and (not (member (strcase name) (mapcar 'strcase (yyc:profiles))))
                                  (yyc:yes (strcat "Start \"" name "\" from a copy of \"" (yyc:profile) "\"?") "Yes")))
    (yyc:add-profile name)
  )
  (yyc:set "Description" (yyc:ask "Short description, e.g. Concourse B - Domestic (04 grid)" (yyc:get "Description" (yyc:profile))))
  (yyc:msg (strcat "Kit folder now: " (yyc:get "Kit" "(not set)") " - pick a new one, or Cancel to keep it."))
  (if (setq kit (yyc:browse-folder (strcat "Kit folder for profile \"" (yyc:profile) "\" (Cancel keeps the current one)")))
    (yyc:set "Kit" kit))
  (foreach kind *yyc-file-kinds* (yyc:choose-file (car kind)))
  (yyc:set "PageSetup" (yyc:ask "Page setup name inside the template" (yyc:get "PageSetup" "YYC - Titleblock")))
  (yyc:set "DwgNo"     (yyc:ask "YYC drawing number" (yyc:get "DwgNo" "24C024")))
  (yyc:set "CTB"       (yyc:ask "Plot style table" (yyc:get "CTB" "YYC_BW_HPv1.ctb")))
  (yyc:set "Fonts"     (yyc:ask "Allowed font files, comma separated" (yyc:get "Fonts" "caa_eng.shx,caa_arch.shx,CAA.SHX")))
  (yyc:set "Scales"    (yyc:ask "Standard viewport scales" (yyc:get "Scales" "1:1,1:2,1:5,1:10,1:20,1:25,1:50,1:75,1:100,1:125,1:200,1:250,1:500,1:1000")))
  (yyc:set "SchedMargin" (yyc:ask "YYCSCHEDULES margin past the viewport frame, in paper mm" (yyc:get "SchedMargin" "5")))
  (yyc:print-settings)
  (yyc:msg "Check or switch profiles any time with YYCPROFILES.")
)

(defun yyc:print-settings ()
  (yyc:msg (strcat "---- YYC CAD Tools - profile \"" (yyc:profile) "\" (others: " (yyc:join (yyc:profiles) ", ") ") ----"))
  (foreach k *yyc-keys* (yyc:msg (strcat "  " k ": " (yyc:get k "(not set)"))))
  
  (princ)
)

;;; ---------------------------------------------------------------------------
;;; Step 2 - Rename exported files
;;; ---------------------------------------------------------------------------

;; Revit names like "Project-Sheet - A102 - Floor Plan" -> "A102"
(defun yyc:sheet-token (base / p)
  (if (setq p (vl-string-search "Sheet - " base)) (setq base (substr base (+ p 9))))
  (if (setq p (vl-string-search " - " base)) (setq base (substr base 1 p)))
  base
)

(defun yyc:clean-name (base dwgno / s)
  (setq s (strcase (yyc:sheet-token base)))
  (foreach ch '("." "-" " " "_")
    (while (vl-string-search ch s) (setq s (vl-string-subst "" ch s)))
  )
  (if (/= (substr s 1 (strlen dwgno)) (strcase dwgno)) (setq s (strcat (strcase dwgno) s)))
  (strcat s ".dwg")
)

(defun c:YYCRENAME ( / folder dwgno files pairs new ok fail opn)
  (setq folder (yyc:browse-folder "Folder of exported DWGs to rename"))
  (if (not folder)
    (yyc:msg "Cancelled.")
    (progn
      (setq dwgno (strcase (yyc:ask "YYC drawing number" (yyc:get "DwgNo" "24C024"))))
      (yyc:set "DwgNo" dwgno)
      (setq files (yyc:dwgs-in folder) pairs '())
      (foreach f files
        (setq new (yyc:clean-name (vl-filename-base f) dwgno))
        (if (/= (strcase f) (strcase new)) (setq pairs (cons (cons f new) pairs)))
      )
      (setq pairs (reverse pairs))
      (if (not pairs)
        (yyc:msg "Every file already matches the naming pattern. Nothing to rename.")
        (progn
          (setq opn (yyc:open-docs))
          (yyc:msg "---- Preview (old  ->  new) ----")
          (foreach p pairs
            (yyc:msg (strcat "  " (car p) "  ->  " (cdr p)
                             (if (member (strcase (strcat folder "\\" (car p))) opn) "   << OPEN IN AUTOCAD - close it first, it will be skipped" ""))))
          (if (vl-some '(lambda (p) (member (strcase (strcat folder "\\" (car p))) opn)) pairs)
            (yyc:msg "Tip: close those drawings and run YYCRENAME from a blank one (NEW)."))
          (yyc:msg "Close these drawings first. If the Revit export wrote views as xrefs, renaming those xref files breaks their paths.")
          (if (yyc:yes (strcat "Rename " (itoa (length pairs)) " file(s)?") "No")
            (progn
              (setq ok 0 fail 0)
              (foreach p pairs
                (if (and (not (findfile (strcat folder "\\" (cdr p))))
                         (vl-file-rename (strcat folder "\\" (car p)) (strcat folder "\\" (cdr p))))
                  (setq ok (1+ ok))
                  (progn (setq fail (1+ fail)) (yyc:msg (strcat "  NOT renamed: " (car p) (if (member (strcase (strcat folder "\\" (car p))) opn) " (open in AutoCAD)" " (open elsewhere, or the new name is taken)"))))
                )
              )
              (yyc:msg (strcat "Renamed " (itoa ok) ", failed " (itoa fail) "."))
            )
            (yyc:msg "Nothing renamed.")
          )
        )
      )
    )
  )
  (princ)
)


;;; ---------------------------------------------------------------------------
;;; Step 3 - Page setup (read from the .dwt with ObjectDBX, copied with CopyFrom)
;;; ---------------------------------------------------------------------------

;; Move everything in a layout so the sheet's lower-left corner sits on 0,0.
;; The corner is taken from the exported titleblock block.
(defun yyc:sheet-to-origin (lay / tb lo hi dx dy mat n)
  (vlax-for o (vla-get-Block lay)
    (if (and (not tb) (= (vla-get-ObjectName o) "AcDbBlockReference") (wcmatch (strcase (vla-get-Name o)) "*TITLEBLOCK*"))
      (setq tb o)))
  (cond
    ((not tb) (yyc:msg (strcat "  " (vla-get-Name lay) ": no titleblock block found - sheet position left as it is.")))
    ((yyc:err-p (vl-catch-all-apply 'vla-GetBoundingBox (list tb 'lo 'hi))) nil)
    (T
     (setq lo (vlax-safearray->list lo) dx (- (car lo)) dy (- (cadr lo)))
     (if (or (> (abs dx) 0.01) (> (abs dy) 0.01))
       (progn
         (setq mat (vlax-tmatrix (list (list 1.0 0.0 0.0 dx) (list 0.0 1.0 0.0 dy) '(0.0 0.0 1.0 0.0) '(0.0 0.0 0.0 1.0))) n 0)
         (vlax-for o (vla-get-Block lay)
           (if (not (yyc:err-p (vl-catch-all-apply 'vla-TransformBy (list o mat)))) (setq n (1+ n))))
         (yyc:msg (strcat "  " (vla-get-Name lay) ": sheet moved " (rtos dx 2 2) ", " (rtos dy 2 2)
                          " mm so its lower-left corner sits on 0,0 (" (itoa n) " objects, viewports included)."))))))
)

(defun yyc:t-pagesetup (doc / dwt ps dbx src pcs pc n)
  (setq dwt (yyc:need-file "Template" "Select the YYC titleblock template" "dwt")
        ps  (yyc:get "PageSetup" "YYC - Titleblock"))
  (if (and dwt (setq dbx (yyc:dbx-open dwt)))
    (progn
      (vlax-for c (vla-get-PlotConfigurations dbx)
        (if (= (strcase (vla-get-Name c)) (strcase ps)) (setq src c))
      )
      (if (not src)
        (yyc:msg (strcat "ERROR: page setup \"" ps "\" is not in " dwt))
        (progn
          (setq pcs (vla-get-PlotConfigurations doc))
          (vl-catch-all-apply '(lambda () (vla-Delete (vla-Item pcs ps))))
          (setq pc (vla-Add pcs ps :vlax-false) n 0)
          (if (yyc:err-p (vl-catch-all-apply 'vla-CopyFrom (list pc src)))
            (yyc:msg "ERROR: could not copy the page setup out of the template."))
          (foreach lay (yyc:paper-layouts doc)
            (if (yyc:err-p (vl-catch-all-apply 'vla-CopyFrom (list lay pc)))
              (yyc:msg (strcat "  FAILED on layout " (vla-get-Name lay)))
              (progn (setq n (1+ n)) (yyc:try 'vla-RefreshPlotDeviceInfo (list lay)) (yyc:sheet-to-origin lay)
                     (yyc:msg (strcat "  " (vla-get-Name lay) ": " (vla-get-CanonicalMediaName lay) ", " (vla-get-StyleSheet lay))))
            )
          )
          (if (yyc:active-p doc) (vl-catch-all-apply 'vla-Regen (list doc acAllViewports)))
          (yyc:msg (strcat "Page setup \"" ps "\" applied to " (itoa n) " layout(s)."))
          (yyc:msg "If the sheet looks unchanged, type RE (REGEN) - the new A0 size shows after a regen.")
        )
      )
      (yyc:dbx-close dbx)
    )
  )
)
(defun c:YYCPAGESETUP () (yyc:run 'yyc:t-pagesetup))

;;; ---------------------------------------------------------------------------
;;; Step 5/6 - YYC grid in and out
;;; ---------------------------------------------------------------------------

(defun yyc:t-gridin (doc / grid ref)
  (setq grid (if (yyc:batch-p) (yyc:need-file "Grid" "Select the grid drawing" "dwg") (yyc:choose-file "Grid")))
  (if grid
    (if (yyc:err-p (setq ref (vl-catch-all-apply 'vla-InsertBlock
                       (list (vla-get-ModelSpace doc) (vlax-3d-point 0 0 0) grid 1.0 1.0 1.0 0.0))))
      (yyc:msg (strcat "ERROR inserting grid: " (vl-catch-all-error-message ref)))
      (progn
        (yyc:set "GridBlock" (vla-get-Name ref))
        (yyc:msg (strcat "Grid inserted in model space at 0,0,0 as block \"" (vla-get-Name ref) "\"."))
        (yyc:msg "Check one known grid intersection with ID first - if it's off, the units didn't match.")
        (yyc:msg "Next: YYCALIGNREC on the first sheet of this terminal (four clicks), or YYCALIGNAPPLY on every other sheet. Then YYCGRIDOUT.")
      )
    )
  )
)
(defun c:YYCGRIDIN () (yyc:run 'yyc:t-gridin))

(defun yyc:t-gridout (doc / name kill)
  (setq name (strcase (yyc:get "GridBlock" (vl-filename-base (yyc:get "Grid" "grid")))))
  (vlax-for o (vla-get-ModelSpace doc)
    (if (and (= (vla-get-ObjectName o) "AcDbBlockReference") (= (strcase (vla-get-Name o)) name))
      (setq kill (cons o kill)))
  )
  (foreach o kill (vl-catch-all-apply 'vla-Delete (list o)))
  (vl-catch-all-apply '(lambda () (vla-Delete (vla-Item (vla-get-Blocks doc) name))))
  (yyc:msg (strcat "Deleted " (itoa (length kill)) " grid block(s) \"" name "\". Now run YYCCLEAN."))
)
(defun c:YYCGRIDOUT () (yyc:run 'yyc:t-gridout))


;;; ---------------------------------------------------------------------------
;;; Step 4 - YYCSCHEDULES: pick a viewport, everything it shows moves to paper
;;; space at the same size and spot (what CHSPACE does), then the empty viewport
;;; can be deleted. One viewport at a time; Enter to finish.
;;; ---------------------------------------------------------------------------

;; viewport geometry from its DXF: frame centre P, frame w/h, view centre C (DCS),
;; target T, twist, scale k (paper units per model unit)
(defun yyc:vp-geom (e / d)
  (setq d (entget e))
  (list (cdr (assoc 10 d)) (cdr (assoc 40 d)) (cdr (assoc 41 d)) (cdr (assoc 12 d))
        (cdr (assoc 17 d)) (cond ((cdr (assoc 51 d))) (0.0)) (/ (cdr (assoc 41 d)) (cdr (assoc 45 d))))
)

;; model point -> DCS of the viewport
(defun yyc:to-dcs (pt tg tw / x y)
  (setq x (- (car pt) (car tg)) y (- (cadr pt) (cadr tg)))
  (list (- (* (cos tw) x) (* (sin tw) y)) (+ (* (sin tw) x) (* (cos tw) y)))
)

;; layers you can't see in this viewport: off, frozen, or frozen in this viewport only
(defun yyc:vp-hidden-layers (doc e / res)
  (vlax-for L (vla-get-Layers doc)
    (if (or (= (vla-get-LayerOn L) :vlax-false) (= (vla-get-Freeze L) :vlax-true))
      (setq res (cons (strcase (vla-get-Name L)) res))))
  (foreach pr (entget e)
    (if (= (car pr) 341) (setq res (cons (strcase (cdr (assoc 2 (entget (cdr pr))))) res))))
  res
)

;; Model objects the viewport shows: on a visible layer and touching the viewport
;; window (crossing, so text that runs a little past the frame comes too), with a
;; small margin past the frame (MARGIN paper mm) for border lines under the edge.
;; Returns (visible locked-count partly-count)
;; What a schedule viewport shows. Small things (text, lines, table cells) are taken like a
;; crossing window, so text hanging past the frame comes too. A BLOCK is only taken when it
;; sits inside the frame (its middle inside, and not bigger than the frame) - a plan block
;; that just overlaps the corner stays in model space. Anything much bigger than the
;; frame (a long plan line passing through) stays too.
(defun yyc:vp-contents (doc e g margin / fp w h vc tg tw k hw hh lo hi in part lock pts ok any hid lay xs ys
                                      ow oh cx cy blk big skip)
  (mapcar 'set '(fp w h vc tg tw k) g)
  (setq hw (/ (+ (/ w 2.0) margin) k) hh (/ (+ (/ h 2.0) margin) k) part 0 lock 0
        hid (yyc:vp-hidden-layers doc e))
  (vlax-for o (vla-get-ModelSpace doc)
    (setq lay (strcase (vla-get-Layer o)))
    (if (and (not (member lay hid))
             (not (yyc:err-p (vl-catch-all-apply 'vla-GetBoundingBox (list o 'lo 'hi)))))
      (progn
        (setq lo (vlax-safearray->list lo) hi (vlax-safearray->list hi))
        (setq pts (mapcar '(lambda (q) (yyc:to-dcs q tg tw))
                          (list lo hi (list (car lo) (cadr hi)) (list (car hi) (cadr lo)))))
        ;; crossing test: the object's extents (in the viewport's own axes) touch the window
        (setq xs (mapcar 'car pts) ys (mapcar 'cadr pts))
        (setq any (and (<= (apply 'min xs) (+ (car vc) hw)) (>= (apply 'max xs) (- (car vc) hw))
                       (<= (apply 'min ys) (+ (cadr vc) hh)) (>= (apply 'max ys) (- (cadr vc) hh))))
        (setq ok (and any (<= (apply 'max (mapcar '(lambda (x) (abs (- x (car vc)))) xs)) hw)
                          (<= (apply 'max (mapcar '(lambda (y) (abs (- y (cadr vc)))) ys)) hh)))
        (setq ow (- (apply 'max xs) (apply 'min xs)) oh (- (apply 'max ys) (apply 'min ys))
              cx (/ (+ (apply 'max xs) (apply 'min xs)) 2.0) cy (/ (+ (apply 'max ys) (apply 'min ys)) 2.0)
              blk (= (vla-get-ObjectName o) "AcDbBlockReference")
              big (or (> ow (* 3.0 hw)) (> oh (* 3.0 hh))))
        (cond
          ((not any) nil)
          ((and (not ok) blk
                (or (> (abs (- cx (car vc))) hw) (> (abs (- cy (cadr vc))) hh) (> ow (* 2.2 hw)) (> oh (* 2.2 hh))))
           (setq skip (cons o skip)))
          ((and (not ok) big) (setq skip (cons o skip)))
          ((= (vla-get-Lock (vla-Item (vla-get-Layers doc) (vla-get-Layer o))) :vlax-true) (setq lock (1+ lock)))
          (T (setq in (cons o in)) (if (not ok) (setq part (1+ part)))))
      )
    )
  )
  (list in lock part skip)
)

(defun yyc:vp-matrix (g / fp w h vc tg tw k kc ks)
  (mapcar 'set '(fp w h vc tg tw k) g)
  (setq kc (* k (cos tw)) ks (* k (sin tw)))
  (vlax-tmatrix
    (list (list kc (- ks) 0.0 (- (car fp) (* k (car vc)) (- (* kc (car tg)) (* ks (cadr tg)))))
          (list ks kc 0.0 (- (cadr fp) (* k (cadr vc)) (+ (* ks (car tg)) (* kc (cadr tg)))))
          (list 0.0 0.0 k 0.0)
          '(0.0 0.0 0.0 1.0)))
)

(defun yyc:type-summary (objs / by)
  (foreach o objs (setq by (yyc:inc (substr (vla-get-ObjectName o) 5) by)))
  (yyc:join (mapcar '(lambda (x) (strcat (itoa (cdr x)) " " (car x))) by) ", ")
)

(defun c:YYCSCHEDULES ( / doc sel e o g res objs arr copies mat main n total)
  (setq doc (yyc:doc) total 0)
  (yyc:msg (strcat "Profile: " (yyc:profile) ". Moves what a viewport shows into paper space, like CHSPACE."))
  (yyc:msg (strcat "Takes everything the viewport touches, like a crossing window (layers on, not frozen there), plus " (yyc:get "SchedMargin" "5") " mm past the frame."))
  (if (= (vla-get-ActiveSpace doc) acModelSpace)
    (yyc:msg "Go to a layout tab first, then run YYCSCHEDULES again.")
    (progn
      (vl-catch-all-apply 'vla-put-MSpace (list doc :vlax-false))
      (setq main (mapcar '(lambda (x) (vla-get-Handle (cdr x))) (yyc:vp-main doc)))
      (while (setq sel (entsel "\nClick a schedule viewport's frame (Enter to finish): "))
        (setq e (car sel) o (vlax-ename->vla-object e))
        (cond
          ((/= (vla-get-ObjectName o) "AcDbViewport") (yyc:msg "That's not a viewport frame - click the frame's edge."))
          ((and (member (vla-get-Handle o) main)
                (not (yyc:yes "This is the biggest viewport - the drawing itself. Move its contents anyway?" "No")))
           (yyc:msg "Left alone."))
          (T
           (setq g (yyc:vp-geom e) res (yyc:vp-contents doc e g (atof (yyc:get "SchedMargin" "5"))) objs (car res))
           (if (not objs)
             (progn
               (yyc:msg "This viewport shows nothing that fits inside it.")
               (if (yyc:yes "Delete this empty viewport?" "Yes") (vl-catch-all-apply 'vla-Delete (list o))))
             (progn
               (yyc:msg (strcat "Found " (itoa (length objs)) " object(s) touching the viewport: " (yyc:type-summary objs)))
               (if (> (caddr res) 0)
                 (yyc:msg (strcat "  " (itoa (caddr res)) " of them run past the viewport edge - included, like a crossing window.")))
               (if (> (cadr res) 0)
                 (yyc:msg (strcat "  " (itoa (cadr res)) " are on locked layers - skipped. Unlock them and run again to include them.")))
               (if (nth 3 res)
                 (yyc:msg (strcat "  Left in model space: " (itoa (length (nth 3 res))) " object(s) that only overlap the frame ("
                                  (yyc:type-summary (nth 3 res)) ") - a block or line from the plan, not part of the schedule.")))
               (if (yyc:yes "Move them to paper space?" "Yes")
                 (progn
                   (vla-StartUndoMark doc)
                   (setq arr (vlax-make-safearray vlax-vbObject (cons 0 (1- (length objs)))))
                   (vlax-safearray-fill arr objs)
                   (setq copies (vl-catch-all-apply 'vla-CopyObjects
                                  (list doc arr (vla-get-Block (vla-get-ActiveLayout doc)))))
                   (if (yyc:err-p copies)
                     (yyc:msg (strcat "ERROR: " (vl-catch-all-error-message copies)))
                     (progn
                       (setq mat (yyc:vp-matrix g) n 0)
                       (foreach c (vlax-safearray->list (vlax-variant-value copies))
                         (if (not (yyc:err-p (vl-catch-all-apply 'vla-TransformBy (list c mat)))) (setq n (1+ n))))
                       (foreach x objs (vl-catch-all-apply 'vla-Delete (list x)))
                       (setq total (+ total n))
                       (yyc:msg (strcat "Moved " (itoa n) " object(s) to paper space."))
                       (if (yyc:yes "Delete the viewport now (it's empty)?" "Yes")
                         (progn (vl-catch-all-apply 'vla-Delete (list o)) (yyc:msg "Viewport deleted. When you're done: YYCVPLAYERS, then YYCCLEAN, to tidy the VIEWPORT# layers."))
                         (yyc:msg "Viewport kept - delete it yourself when you're ready."))))
                   (vla-EndUndoMark doc)
                 )
               )
             )
           )
          )
        )
      )
      (yyc:msg (strcat "Done. " (itoa total) " object(s) moved in total. Each viewport is one undo step."))
    )
  )
  (princ)
)

;;; ---------------------------------------------------------------------------
;;; Step 5 - Record the MOVE/ALIGN once, apply it to the other drawings
;;; Same maths as ALIGN with two point pairs and "No" to scaling: the first pair
;;; sets where things land, the second pair sets the rotation. Model space only.
;;; A profile either has its own recording, or uses another profile's
;;; ("AlignFrom") - so one recording can serve the whole project, or each area
;;; can have its own.
;;; ---------------------------------------------------------------------------

(defun yyc:pt->str (p) (strcat (rtos (car p) 2 6) "," (rtos (cadr p) 2 6)))
(defun yyc:str->pt (s / l) (if (and s (= (length (setq l (yyc:split s ","))) 2)) (list (atof (car l)) (atof (cadr l)) 0.0)))

;; profile whose recording this profile uses (itself unless told otherwise)
(defun yyc:align-src (profile / f)
  (setq f (yyc:get-in profile "AlignFrom" nil))
  (if (and f (/= (strcase f) (strcase profile)) (yyc:get-in f "AlignS1" nil)) f profile)
)

(defun yyc:align-pts (profile / src)
  (setq src (yyc:align-src profile))
  (mapcar '(lambda (k) (yyc:str->pt (yyc:get-in src k nil))) '("AlignS1" "AlignD1" "AlignS2" "AlignD2"))
)

;; 4x4 matrix: rotate by (angle d1->d2 minus angle s1->s2), then s1 lands on d1
(defun yyc:align-matrix (s1 d1 s2 d2 / th c sn tx ty)
  (setq th (- (angle d1 d2) (angle s1 s2)) c (cos th) sn (sin th))
  (setq tx (- (car d1) (- (* c (car s1)) (* sn (cadr s1))))
        ty (- (cadr d1) (+ (* sn (car s1)) (* c (cadr s1)))))
  (vlax-tmatrix (list (list c (- sn) 0.0 tx) (list sn c 0.0 ty) '(0.0 0.0 1.0 0.0) '(0.0 0.0 0.0 1.0)))
)

(defun yyc:deg (r) (* 180.0 (/ r pi)))

(defun yyc:align-describe (s1 d1 s2 d2)
  (strcat "rotation " (rtos (yyc:deg (- (angle d1 d2) (angle s1 s2))) 2 4) " deg, first point moves "
          (rtos (distance s1 d1) 2 0) " units")
)

;; marker stored in the drawing so a move is never applied twice
(defun yyc:aligned-mark (doc / si v)
  (setq si (vla-get-SummaryInfo doc))
  (if (yyc:err-p (vl-catch-all-apply 'vla-GetCustomByKey (list si "YYC_Aligned" 'v))) nil v)
)
(defun yyc:set-aligned-mark (doc text / si)
  (setq si (vla-get-SummaryInfo doc))
  (if (yyc:aligned-mark doc)
    (vl-catch-all-apply 'vla-SetCustomByKey (list si "YYC_Aligned" text))
    (vl-catch-all-apply 'vla-AddCustomInfo (list si "YYC_Aligned" text)))
)

;; move every model-space object except the grid block; returns (moved . lockedLayers)
(defun yyc:transform-model (doc mat / grid n locked)
  (setq grid (strcase (yyc:get "GridBlock" "")) n 0)
  (vlax-for o (vla-get-ModelSpace doc)
    (if (not (and (= (vla-get-ObjectName o) "AcDbBlockReference") (= (strcase (vla-get-Name o)) grid)))
      (if (yyc:err-p (vl-catch-all-apply 'vla-TransformBy (list o mat)))
        (if (not (member (vla-get-Layer o) locked)) (setq locked (cons (vla-get-Layer o) locked)))
        (setq n (1+ n))))
  )
  (cons n locked)
)

(defun yyc:report-move (res)
  (yyc:msg (strcat "Moved " (itoa (car res)) " model-space object(s). Paper space was not touched."))
  (if (cdr res) (yyc:msg (strcat "NOT moved - on locked layers (unlock, then run again): " (yyc:join (cdr res) ", "))))
)

;; Give a viewport a UCS that lines up with its view, the way UCS OB + PLAN does by
;; hand: crosshairs, ortho and coordinates inside the viewport then follow the sheet.
;; It steps INTO the viewport (MSPACE + CVPORT = that viewport's ID), sets UCSVP 1
;; and UCS Z, then goes back to paper space and keeps the sheet's own UCS at World.
;; Needs the drawing on screen with that viewport's layout active.
(defun yyc:vp-set-ucs (doc vp / ang e id echo ok)
  (setq ang (- (* 2 pi) (vla-get-TwistAngle vp)))
  (if (and (yyc:active-p doc)
           (setq e (yyc:try 'vlax-vla-object->ename (list vp)))
           (setq id (cdr (assoc 69 (entget e))))
           (> id 1))
    (progn
      (setq echo (getvar "CMDECHO")) (setvar "CMDECHO" 0)
      (vl-catch-all-apply 'vla-put-ViewportOn (list vp :vlax-true))
      (command "_.MSPACE")
      (if (not (yyc:err-p (vl-catch-all-apply 'setvar (list "CVPORT" id))))
        (progn
          (setvar "UCSVP" 1)
          (command "_.UCS" "_World")
          (command "_.UCS" "_Z" (rtos (* 180.0 (/ ang pi)) 2 8))
          (setq ok (= (getvar "CVPORT") id))))
      (command "_.PSPACE")
      (command "_.UCS" "_World")          ; the sheet itself stays square
      (setvar "CMDECHO" echo)
    )
  )
  ok
)

;; Keep each viewport showing the same part of the plan after the move, turned
;; so the sheet reads exactly as exported (project north up). The model moved by
;; rotation TH and shift (TX,TY), so each viewport's target moves the same way and
;; its view twist turns by -TH. Scale and view centre stay as they were.
(defun yyc:vp-follow (doc vps s1 d1 s2 d2 / th c sn tx ty n tg nt lk vp)
  (setq th (- (angle d1 d2) (angle s1 s2)) c (cos th) sn (sin th))
  (setq tx (- (car d1) (- (* c (car s1)) (* sn (cadr s1))))
        ty (- (cadr d1) (+ (* sn (car s1)) (* c (cadr s1)))) n 0)
  (foreach pr vps
    (setq vp (cdr pr) lk (vla-get-DisplayLocked vp))
    (vl-catch-all-apply 'vla-put-DisplayLocked (list vp :vlax-false))
    (setq tg (vlax-safearray->list (vlax-variant-value (vla-get-Target vp))))
    (setq nt (list (+ tx (- (* c (car tg)) (* sn (cadr tg)))) (+ ty (+ (* sn (car tg)) (* c (cadr tg)))) (caddr tg)))
    (if (and (not (yyc:err-p (vl-catch-all-apply 'vla-put-Target (list vp (vlax-3d-point nt)))))
             (not (yyc:err-p (vl-catch-all-apply 'vla-put-TwistAngle (list vp (rem (+ (- (vla-get-TwistAngle vp) th) (* 4 pi)) (* 2 pi)))))))
      (progn
        (setq n (1+ n))
        (if (not (yyc:vp-set-ucs doc vp))
          (yyc:msg (strcat "  View turned in " (car pr) ", but its UCS could not be set - run YYCVPUCS with the sheet open."))))
      (yyc:msg (strcat "  Could not turn the viewport in " (car pr))))
    (vl-catch-all-apply 'vla-put-DisplayLocked (list vp lk))
  )
  (yyc:set-mark doc "YYC_VPFollowed" (yyc:stamp))
  (yyc:msg (strcat (itoa n) " viewport(s) now show the same area as the export, turned to project north (twist "
                   (rtos (yyc:deg (- th)) 2 3) " deg), with a matching UCS inside each one. Check one; if it is rotated the wrong way, tell us."))
  n
)

;; the drawing viewport of each layout = the biggest one
(defun yyc:vp-main (doc / best res)
  (foreach lay (yyc:paper-layouts doc)
    (setq best nil)
    (foreach pr (yyc:vp-list doc)
      (if (and (= (car pr) (vla-get-Name lay))
               (or (not best) (> (* (vla-get-Width (cdr pr)) (vla-get-Height (cdr pr)))
                                 (* (vla-get-Width (cdr best)) (vla-get-Height (cdr best))))))
        (setq best pr)))
    (if best (setq res (cons best res))))
  (reverse res)
)

;; which viewports to turn: "Main" (drawing viewport only, the default) or "All"
(defun yyc:vp-choice (doc ask / k all main)
  (setq all (yyc:vp-list doc) main (yyc:vp-main doc))
  (if (and ask (> (length all) (length main)))
    (progn
      (yyc:msg (strcat (itoa (length all)) " viewports. The drawing viewport (biggest) in each layout:"))
      (foreach pr main (yyc:msg (strcat "    " (car pr) ": " (rtos (vla-get-Width (cdr pr)) 2 0) " x " (rtos (vla-get-Height (cdr pr)) 2 0) ", scale " (yyc:scale-text (cdr pr)))))
      (initget "Main All None")
      (setq k (getkword "\nTurn which viewports? [Main = drawing viewport only/All/None] <Main>: "))
      (if (not k) (setq k "Main"))
      (yyc:set "VPFollow" k))
    (setq k (yyc:get "VPFollow" "Main")))
  (cond ((= k "All") all) ((= k "None") nil) (T main))
)

;; Turn the chosen viewports. Each sheet is made the active layout first, so its
;; viewports are live and the turn sticks (from the Model tab they aren't), then
;; you're put back on the tab you started from.
(defun yyc:vp-follow-run (doc ask pts / start mode lays lname vps all n)
  (setq start (vla-get-ActiveLayout doc) n 0)
  ;; ask once (Main / All / None), using what the drawing has
  (setq mode (if (yyc:vp-choice doc ask) (yyc:get "VPFollow" "Main") "None"))
  (if (= mode "None") (setq mode (yyc:get "VPFollow" "Main")))
  (if (/= (yyc:get "VPFollow" "Main") "None")
    (foreach lay (yyc:paper-layouts doc)
      (setq lname (vla-get-Name lay))
      (vl-catch-all-apply 'vla-put-ActiveLayout (list doc lay))
      (vl-catch-all-apply 'vla-put-MSpace (list doc :vlax-false))
      (setq all (vl-remove-if-not '(lambda (x) (= (car x) lname)) (yyc:vp-list doc)))
      (setq vps (if (= (yyc:get "VPFollow" "Main") "All") all
                  (vl-remove-if-not '(lambda (x) (= (car x) lname)) (yyc:vp-main doc))))
      (if vps (setq n (+ n (apply 'yyc:vp-follow (append (list doc vps) pts)))))
      (foreach pr all
        (if (not (member (vla-get-Handle (cdr pr)) (mapcar '(lambda (x) (vla-get-Handle (cdr x))) vps)))
          (yyc:msg (strcat "  Left as it was: viewport in " (car pr) " (" (rtos (vla-get-Width (cdr pr)) 2 0) " x "
                           (rtos (vla-get-Height (cdr pr)) 2 0) ") - it still looks at the old location. Delete it if it's empty."))))
    )
  )
  (vl-catch-all-apply 'vla-put-ActiveLayout (list doc start))
  (vl-catch-all-apply 'vla-Regen (list doc acAllViewports))
  (yyc:msg (strcat "Viewports turned: " (itoa n) ". You're back on the " (vla-get-Name start) " tab."))
)

(defun yyc:get-mark (doc key / si v)
  (setq si (vla-get-SummaryInfo doc))
  (if (yyc:err-p (vl-catch-all-apply 'vla-GetCustomByKey (list si key 'v))) nil v)
)
(defun yyc:set-mark (doc key text / si)
  (setq si (vla-get-SummaryInfo doc))
  (if (yyc:get-mark doc key)
    (vl-catch-all-apply 'vla-SetCustomByKey (list si key text))
    (vl-catch-all-apply 'vla-AddCustomInfo (list si key text)))
)

(defun yyc:t-vpfollow (doc / pts)
  (setq pts (yyc:align-pts (yyc:profile)))
  (cond
    ((member nil pts) (yyc:msg "No recorded move in this profile - run YYCALIGNREC first."))
    ((yyc:get-mark doc "YYC_VPFollowed") (yyc:msg (strcat "SKIPPED - viewports already turned: " (yyc:get-mark doc "YYC_VPFollowed"))))
    ((not (yyc:aligned-mark doc)) (yyc:msg "This drawing hasn't been moved yet - run YYCALIGNAPPLY first."))
    (T (yyc:vp-follow-run doc (not (yyc:batch-p)) pts))
  )
)

;; Fix drawings turned before the UCS was set: give each drawing viewport a UCS
;; matching its current view twist.
(defun yyc:t-vpucs (doc / start n mode vps)
  (setq start (vla-get-ActiveLayout doc) n 0)
  (foreach lay (yyc:paper-layouts doc)
    (vl-catch-all-apply 'vla-put-ActiveLayout (list doc lay))
    (vl-catch-all-apply 'vla-put-MSpace (list doc :vlax-false))
    (setq vps (vl-remove-if-not '(lambda (x) (= (car x) (vla-get-Name lay)))
                (if (= (yyc:get "VPFollow" "Main") "All") (yyc:vp-list doc) (yyc:vp-main doc))))
    (foreach pr vps
      (if (yyc:vp-set-ucs doc (cdr pr))
        (progn (setq n (1+ n))
               (yyc:msg (strcat "  " (car pr) ": UCS set to the view (" (rtos (yyc:deg (- (* 2 pi) (vla-get-TwistAngle (cdr pr)))) 2 3) " deg)")))
        (yyc:msg (strcat "  " (car pr) ": could not set the UCS")))))
  (vl-catch-all-apply 'vla-put-ActiveLayout (list doc start))
  (yyc:msg (strcat (itoa n) " viewport(s) now have a UCS matching their view. The sheet's own UCS is World."))
)
(defun c:YYCVPUCS () (yyc:run 'yyc:t-vpucs))



(defun c:YYCALIGNREC ( / doc s1 d1 s2 d2 ratio res ps ans picks)
  (setq doc (yyc:doc))
  (yyc:msg (strcat "Profile: " (yyc:profile) ". Records a move, the same way as ALIGN with two pairs and No to scale."))
  (if (yyc:aligned-mark doc) (yyc:msg (strcat "NOTE: this drawing is already marked as moved: " (yyc:aligned-mark doc))))
  (if (/= (vla-get-ActiveSpace doc) acModelSpace)
    (vla-put-ActiveLayout doc (vla-Item (vla-get-Layouts doc) "Model")))
  (if (and (setq s1 (getpoint "\nFirst SOURCE point (on the export): "))
           (setq d1 (getpoint s1 "\nFirst DESTINATION point (same spot on the YYC grid): "))
           (setq s2 (getpoint "\nSecond SOURCE point (far from the first): "))
           (setq d2 (getpoint s2 "\nSecond DESTINATION point: ")))
    (progn
      (setq s1 (trans s1 1 0) d1 (trans d1 1 0) s2 (trans s2 1 0) d2 (trans d2 1 0))
      (setq ratio (/ (distance d1 d2) (max (distance s1 s2) 1e-9)))
      (yyc:msg (yyc:align-describe s1 d1 s2 d2))
      (if (> (abs (- ratio 1.0)) 0.001)
        (yyc:msg (strcat "WARNING: the two pairs are different lengths (" (rtos (* 100 (- ratio 1)) 2 2)
                         "%). No scaling is applied - check you picked matching points.")))
      (if (yyc:yes "Move model space now and save this to the profile?" "Yes")
        (progn
          (vla-StartUndoMark doc)
          (setq res (yyc:transform-model doc (yyc:align-matrix s1 d1 s2 d2)))
          (vla-EndUndoMark doc)
          (yyc:report-move res)
          (yyc:set "AlignS1" (yyc:pt->str s1)) (yyc:set "AlignD1" (yyc:pt->str d1))
          (yyc:set "AlignS2" (yyc:pt->str s2)) (yyc:set "AlignD2" (yyc:pt->str d2))
          (yyc:set "AlignDate" (yyc:stamp)) (yyc:set "AlignFrom" "")
          (yyc:set-aligned-mark doc (strcat (yyc:profile) " " (yyc:stamp)))
          (yyc:msg (strcat "Saved to profile \"" (yyc:profile) "\". Other drawings: YYCALIGNAPPLY, or Align in YYCBATCH."))
          (yyc:msg "Now turning the drawing viewport so the sheet still shows the same plan, project north up.")
          (yyc:vp-follow-run doc T (list s1 d1 s2 d2))
          ;; share with other profiles
          (setq ps (vl-remove-if '(lambda (x) (= (strcase x) (strcase (yyc:profile)))) (yyc:profiles)))
          (if ps
            (progn
              (yyc:msg "If the other areas export to the same spot, they can use this same move:")
              (setq picks 0)
              (foreach x ps (setq picks (1+ picks)) (yyc:msg (strcat "    " (itoa picks) ". " x)))
              (setq ans (getstring "\nNumbers separated by commas, A for all, Enter for none: "))
              (foreach x (cond ((= (strcase ans) "A") ps)
                               ((= ans "") nil)
                               (T (mapcar '(lambda (i) (if (> (atoi i) 0) (nth (1- (atoi i)) ps))) (yyc:split ans ","))))
                (if x (progn (yyc:set-in x "AlignFrom" (yyc:profile)) (yyc:msg (strcat "  \"" x "\" now uses this move."))))
              )
            )
          )
        )
        (yyc:msg "Nothing moved, nothing saved.")
      )
    )
    (yyc:msg "Cancelled.")
  )
  (princ)
)

(defun yyc:t-alignapply (doc / pts src mark res)
  (setq src (yyc:align-src (yyc:profile)) pts (yyc:align-pts (yyc:profile)) mark (yyc:aligned-mark doc))
  (cond
    ((member nil pts) (yyc:msg (strcat "No recorded move for profile \"" (yyc:profile) "\". Run YYCALIGNREC once for this profile (four clicks), then YYCALIGNAPPLY.")))
    ((and mark (or (yyc:batch-p) (not (yyc:yes (strcat "Already moved (" mark "). Move it AGAIN?") "No"))))
     (yyc:msg (strcat "SKIPPED - already moved: " mark)))
    (T
     (yyc:msg (strcat "Applying the move recorded in \"" src "\" (" (yyc:get-in src "AlignDate" "?") "): " (apply 'yyc:align-describe pts)))
     (setq res (yyc:transform-model doc (apply 'yyc:align-matrix pts)))
     (yyc:report-move res)
     (yyc:set-aligned-mark doc (strcat src " " (yyc:stamp)))
     (if (not (yyc:get-mark doc "YYC_VPFollowed")) (yyc:vp-follow-run doc nil pts))
     (yyc:msg (strcat "Check: ID a known grid intersection. The recorded first point should now read "
                      (rtos (car (nth 1 pts)) 2 0) "," (rtos (cadr (nth 1 pts)) 2 0) ". Or YYCGRIDIN and look."))
    )
  )
)
(defun c:YYCALIGNAPPLY () (yyc:run 'yyc:t-alignapply))

;; let the current profile use another profile's recorded move
(defun c:YYCALIGNUSE ( / me ps i ans src)
  (setq me (yyc:profile) src (yyc:align-src me) i 0)
  (setq ps (vl-remove-if-not '(lambda (x) (and (/= (strcase x) (strcase me)) (yyc:get-in x "AlignS1" nil))) (yyc:profiles)))
  (yyc:msg (strcat "Profile \"" me "\" uses: "
                   (cond ((/= (strcase src) (strcase me)) (strcat "the move recorded in \"" src "\""))
                         ((yyc:get-in me "AlignS1" nil) "its own recorded move")
                         (T "no move yet"))))
  (if (not ps)
    (progn
      (yyc:msg "No OTHER profile has a recorded move to borrow.")
      (yyc:msg "YYCALIGNUSE is for a second profile (another area that exports to the same spot):")
      (yyc:msg "  YYCPROFILES -> pick or create that profile -> Use, then run YYCALIGNUSE and pick this one.")
      (yyc:msg "Same profile, more sheets? You don't need it - just run YYCALIGNAPPLY."))
    (progn
      (yyc:msg (strcat "Borrow a recorded move for \"" me "\":"))
      (foreach x ps (setq i (1+ i)) (yyc:msg (strcat "    " (itoa i) ". " x "  (recorded " (yyc:get-in x "AlignDate" "?") ")")))
      (if (yyc:get-in me "AlignS1" nil) (yyc:msg "    0. go back to this profile's own recording"))
      (setq ans (getstring "\nNumber (Enter to cancel): "))
      (cond
        ((= ans "") (yyc:msg "Nothing changed."))
        ((and (= ans "0") (yyc:get-in me "AlignS1" nil)) (yyc:set "AlignFrom" "") (yyc:msg (strcat "\"" me "\" uses its own recording again.")))
        ((and (> (atoi ans) 0) (<= (atoi ans) (length ps)))
         (yyc:set "AlignFrom" (nth (1- (atoi ans)) ps))
         (yyc:msg (strcat "\"" me "\" now uses the move from \"" (nth (1- (atoi ans)) ps) "\". YYCALIGNAPPLY will use it.")))
        (T (yyc:msg "Not a choice - nothing changed.")))
    )
  )
  (princ)
)

;;; ---------------------------------------------------------------------------
;;; Profiles window - YYCPROFILES (YYCSETUP opens it too): see, edit in place, add, remove, use
;;; ---------------------------------------------------------------------------

(defun yyc:short (path n)
  (cond ((not path) "-") ((<= (strlen path) n) path) (T (strcat "..." (substr path (- (strlen path) n -4)))))
)
(defun yyc:fname (key p / v) (setq v (yyc:get-in p key nil)) (if v (strcat (vl-filename-base v) (cond ((vl-filename-extension v)) (""))) "- not set -"))

(defun yyc:move-text (p / src)
  (setq src (yyc:align-src p))
  (cond ((/= (strcase src) (strcase p)) (strcat "uses the move recorded in \"" src "\""))
        ((yyc:get-in p "AlignS1" nil) (strcat "recorded " (yyc:get-in p "AlignDate" "?") " - " (apply 'yyc:align-describe (yyc:align-pts p))))
        (T "none recorded yet (YYCALIGNREC)"))
)

;; what the window lets you type, and the files it lets you pick
(setq *yyc-ptext*
  '(("Description" "Description" "")
    ("DwgNo"       "YYC drawing no." "24C024")))
;; set once, rarely changed - behind "More settings..."
(setq *yyc-padv*
  '(("PageSetup"   "Page setup name" "YYC - Titleblock")
    ("CTB"         "Plot style table" "YYC_BW_HPv1.ctb")
    ("Scales"      "Viewport scales" "1:1,1:2,1:5,1:10,1:20,1:25,1:50,1:75,1:100,1:125,1:200,1:250,1:500,1:1000")
    ("SchedMargin" "Schedule margin (mm)" "5")
    ("Fonts"       "Allowed fonts" "caa_eng.shx,caa_arch.shx,CAA.SHX")))
;; only the files you need. Layer Reference
;; (only used when there's no standard CSV) and the old layer map CSV still work
;; if set, but aren't shown.
(setq *yyc-pfiles*
  '(("Kit" "Kit folder") ("Grid" "Grid drawing") ("Template" "Template (.dwt)")
    ("StdCSV" "Standard layers CSV") ("LayerBook" "Layer Mapper workbook") ("QABook" "QA Report template") ("FDTemplate" "File Description (.docx)") ("LinFile" "Linetype file (CAA.LIN)")))

(defun yyc:pd-sel () (nth (atoi *yyc-psel*) *yyc-plist*))

(defun yyc:pd-file-text (p k / v)
  (setq v (yyc:get-in p k nil))
  (cond ((not v) "- not set -")
        ((= k "Kit") (strcat (yyc:short v 46) (if (vl-file-directory-p v) "" "   (MISSING)")))
        (T (strcat (vl-filename-base v) (cond ((vl-filename-extension v)) ("")) (if (findfile v) "" "   (MISSING)"))))
)

(defun yyc:pd-list ( / i)
  (start_list "list")
  (foreach x *yyc-plist* (add_list (if (= (strcase x) (strcase (yyc:profile))) (strcat "* " x) (strcat "  " x))))
  (end_list)
  (set_tile "list" *yyc-psel*)
)

(defun yyc:pd-fill (p / users lm)
  ;; the standard list is not a layer map - picked by mistake in older versions
  (if (and (setq lm (yyc:get-in p "LayerMap" nil))
           (or (= (strcase lm) (strcase (yyc:get-in p "StdCSV" ""))) (wcmatch (strcase lm) "*STANDARD*LAYER*")))
    (yyc:set-in p "LayerMap" ""))
  (yyc:pd-msrc-fill p)
  (set_tile "box" (strcat "Settings for \"" p "\"" (if (= (strcase p) (strcase (yyc:profile))) "  (in use)" "")))
  (foreach t1 *yyc-ptext* (set_tile (strcat "e_" (car t1)) (yyc:get-in p (car t1) (caddr t1))))
  (foreach f1 *yyc-pfiles* (set_tile (strcat "f_" (car f1)) (yyc:pd-file-text p (car f1))))
  (setq users (vl-remove-if-not '(lambda (x) (and (/= x p) (= (strcase (yyc:align-src x)) (strcase p)))) *yyc-plist*))
  (set_tile "p7" (strcat "Move:  " (yyc:move-text p)))
  (if users (set_tile "p7" (strcat "Move:  " (yyc:move-text p) "   (also used by " (yyc:join users ", ") ")")))
)

;; write what's typed in the boxes into profile p
;; Grid move: the profile's own recording, or another profile's
(defun yyc:pd-msrc-fill (p)
  (setq *yyc-msrc* (cons nil (vl-remove-if-not '(lambda (x) (and (/= (strcase x) (strcase p)) (yyc:get-in x "AlignS1" nil))) *yyc-plist*)))
  (start_list "msrc")
  (add_list (if (yyc:get-in p "AlignS1" nil) "Its own recording" "Its own recording (none yet - run YYCALIGNREC)"))
  (foreach x (cdr *yyc-msrc*) (add_list (strcat "Same as profile \"" x "\"")))
  (end_list)
  (set_tile "msrc" (itoa (cond ((vl-position (strcase (yyc:align-src p)) (mapcar '(lambda (x) (if x (strcase x))) *yyc-msrc*))) (0))))
)
(defun yyc:pd-msrc-set (v / p q)
  (setq p (yyc:pd-sel) q (nth (atoi v) *yyc-msrc*))
  (yyc:pd-store p)
  (if q
    (progn
      (yyc:set-in p "AlignFrom" q)
      ;; never borrow in a circle: q goes back to its own recording
      (if (= (strcase (yyc:get-in q "AlignFrom" "")) (strcase p)) (yyc:set-in q "AlignFrom" "")))
    (yyc:set-in p "AlignFrom" ""))
  (yyc:pd-fill p)
  (set_tile "msg" (if q (strcat "\"" p "\" now uses the grid move recorded in \"" q "\".") (strcat "\"" p "\" uses its own grid move.")))
)
(defun yyc:pd-clear (k / p)
  (setq p (yyc:pd-sel))
  (yyc:pd-store p)
  (yyc:set-in p k "")
  (yyc:pd-fill p)
  (set_tile "msg" (strcat (cadr (assoc k *yyc-pfiles*)) " cleared."))
)

(defun yyc:pd-adv ( / p r)
  (setq p (yyc:pd-sel))
  (yyc:pd-store p)
  (if (new_dialog "yycadv" *yyc-pdid*)
    (progn
      (foreach t1 *yyc-padv* (set_tile (strcat "a_" (car t1)) (yyc:get-in p (car t1) (caddr t1))))
      (action_tile "accept" "(foreach t1 *yyc-padv* (yyc:set-in (yyc:pd-sel) (car t1) (yyc:trim (get_tile (strcat \"a_\" (car t1)))))) (done_dialog 1)")
      (setq r (start_dialog))
      (set_tile "msg" (if (= r 1) "More settings saved." ""))))
)

(defun yyc:pd-store (p)
  (if p (foreach t1 *yyc-ptext* (yyc:set-in p (car t1) (yyc:trim (get_tile (strcat "e_" (car t1)))))))
)

(defun yyc:pd-add ( / name old)
  (setq name (yyc:trim (get_tile "newname")) old (yyc:pd-sel))
  (cond
    ((= name "") (set_tile "msg" "Type a name in the New box first (e.g. ITB, 24C024 Area 4)."))
    ((member (strcase name) (mapcar 'strcase *yyc-plist*)) (set_tile "msg" (strcat "\"" name "\" already exists.")))
    ((vl-string-search "|" name) (set_tile "msg" "A profile name can't contain |."))
    (T
     (yyc:pd-store old)
     (foreach k *yyc-keys* (if (not (yyc:get-in name k nil)) (yyc:set-in name k (yyc:get-in old k nil))))
     (yyc:set-in name "Description" "")
     (yyc:add-profile name)
     (setq *yyc-plist* (yyc:profiles) *yyc-psel* (itoa (vl-position name *yyc-plist*)))
     (yyc:pd-list) (yyc:pd-fill name) (set_tile "newname" "")
     (set_tile "msg" (strcat "Added \"" name "\" as a copy of \"" old "\". Change what differs (usually the Grid and Description), then Save.")))
  )
)

(defun yyc:pd-del ( / p)
  (setq p (yyc:pd-sel))
  (cond
    ((= (strcase p) (strcase (yyc:profile))) (set_tile "msg" "That's the profile in use - Use another one first, then remove this one."))
    ((/= *yyc-pdel* p) (setq *yyc-pdel* p) (set_tile "msg" (strcat "Click Remove again to take \"" p "\" off the list.")))
    (T
     (setenv "YYC_Profiles" (yyc:join (vl-remove p *yyc-plist*) "|"))
     (setq *yyc-plist* (yyc:profiles) *yyc-psel* "0" *yyc-pdel* nil)
     (yyc:pd-list) (yyc:pd-fill (yyc:pd-sel))
     (set_tile "msg" (strcat "Removed \"" p "\".")))
  )
)

;; fill every empty or missing file from the profile's kit folder (first match)
(defun yyc:pd-auto ( / p kit n v hit)
  (setq p (yyc:pd-sel) kit (yyc:get-in p "Kit" nil) n 0)
  (yyc:pd-store p)
  (if (not (and kit (vl-file-directory-p kit)))
    (set_tile "msg" "Set the Kit folder first (the ... button next to it).")
    (progn
      (foreach kind *yyc-file-kinds*
        (setq v (yyc:get-in p (car kind) nil))
        (if (and (not (and v (findfile v))) (assoc (car kind) *yyc-pfiles*)
                 (setq hit (car (yyc:find-all kit (nth 2 kind) 4))))
          (progn (yyc:set-in p (car kind) hit) (setq n (1+ n)))))
      (yyc:pd-fill p)
      (set_tile "msg" (strcat "Filled " (itoa n) " file(s) from the kit folder. Check the Grid is the right one.")))
  )
)

(defun yyc:pd-dcl (path / f)
  (setq f (open path "w"))
  (foreach x (append
    (list
    "yycprof : dialog { label = \"YYC profiles - pick one, change it here, Save\";"
    "  : row {"
    "    : column {"
    "      : list_box { key = \"list\"; label = \"Profiles  (* = in use)\"; width = 28; height = 20; }"
    "      : edit_box { key = \"newname\"; label = \"New:\"; edit_width = 18; }"
    "      : button { key = \"add\"; label = \"Add (copies the selected one)\"; }"
    "      : button { key = \"del\"; label = \"Remove selected\"; }"
    "    }"
    "    : boxed_column { key = \"box\"; label = \"Settings\";")
    (mapcar '(lambda (t1) (strcat "      : edit_box { key = \"e_" (car t1) "\"; label = \"" (cadr t1) "\"; edit_width = 58; }")) *yyc-ptext*)
    (list "      : spacer { height = 0.5; }")
    (mapcar '(lambda (f1) (strcat "      : row { : text { label = \"" (cadr f1) "\"; width = 24; fixed_width = true; }"
                                  " : text { key = \"f_" (car f1) "\"; width = 50; }"
                                  " : button { key = \"b_" (car f1) "\"; label = \"...\"; width = 5; fixed_width = true; } }")) *yyc-pfiles*)
    (list
    "      : row { : button { key = \"auto\"; label = \"Fill empty files from the kit folder\"; fixed_width = true; }"
    "              : button { key = \"adv\"; label = \"More settings...\"; fixed_width = true; } }"
    "      : popup_list { key = \"msrc\"; label = \"Grid move\"; edit_width = 50; }"
    "      : text { key = \"p7\"; width = 82; }"
    "    }"
    "  }"
    "  : text { key = \"msg\"; width = 112; }"
    "  : row { : button { key = \"save\"; label = \"Save\"; width = 14; }"
    "          : button { key = \"use\"; label = \"Save + use this profile\"; is_default = true; }"
    "          : button { key = \"cancel\"; label = \"Close\"; is_cancel = true; width = 14; } }"
    "}"
    "yycadv : dialog { label = \"More settings - usually left as they are\";")
    (mapcar '(lambda (t1) (strcat "  : edit_box { key = \"a_" (car t1) "\"; label = \"" (cadr t1) "\"; edit_width = 60; }")) *yyc-padv*)
    (list "  ok_cancel;" "}")) (write-line x f))
  (close f)
)

(defun c:YYCPROFILES ( / dcl id res go i k kind p v)
  (setq *yyc-plist* (yyc:profiles) *yyc-psel* "0" *yyc-pdel* nil go T i 0)
  (foreach x *yyc-plist* (if (= (strcase x) (strcase (yyc:profile))) (setq *yyc-psel* (itoa i))) (setq i (1+ i)))
  (setq dcl (vl-filename-mktemp "yycprof.dcl"))
  (yyc:pd-dcl dcl)
  (setq id (load_dialog dcl) *yyc-pdid* id)
  (while go
    (if (and (>= id 0) (new_dialog "yycprof" id))
      (progn
        (yyc:pd-list)
        (yyc:pd-fill (yyc:pd-sel))
        (action_tile "list" "(yyc:pd-store (yyc:pd-sel)) (setq *yyc-psel* $value *yyc-pdel* nil) (yyc:pd-fill (yyc:pd-sel)) (set_tile \"msg\" \"\")")
        (action_tile "add"  "(yyc:pd-add)")
        (action_tile "del"  "(yyc:pd-del)")
        (action_tile "auto" "(yyc:pd-auto)")
        (action_tile "adv"  "(yyc:pd-adv)")
        (action_tile "msrc" "(yyc:pd-msrc-set $value)")
        (action_tile "save" "(yyc:pd-store (yyc:pd-sel)) (set_tile \"msg\" (strcat \"Saved \\\"\" (yyc:pd-sel) \"\\\".\"))")
        (action_tile "use"  "(yyc:pd-store (yyc:pd-sel)) (done_dialog 2)")
        (setq i 0)
        (foreach f1 *yyc-pfiles*
          (action_tile (strcat "b_" (car f1)) (strcat "(yyc:pd-store (yyc:pd-sel)) (done_dialog " (itoa (+ 10 i)) ")"))
          (setq i (1+ i)))
        (setq res (start_dialog) p (yyc:pd-sel))
        (cond
          ((= res 2) (yyc:switch-profile p nil) (yyc:try 'yyc:ensure-support nil) (yyc:msg (strcat "Saved and now using profile \"" p "\".")) (setq go nil))
          ((>= res 10)
           (setq k (car (nth (- res 10) *yyc-pfiles*)))
           (if (= k "Kit")
             (setq v (yyc:browse-folder (strcat "Kit folder for profile \"" p "\"")))
             (setq kind (assoc k *yyc-file-kinds*)
                   v (getfiled (nth 1 kind) (if (yyc:get-in p "Kit" nil) (strcat (yyc:get-in p "Kit" "") "\\") "") (nth 3 kind) 0)))
           (if v (yyc:set-in p k v)))   ; then the window opens again
          (T (setq go nil))
        )
      )
      (progn (yyc:msg "Could not open the profiles window.") (setq go nil))
    )
  )
  (if (>= id 0) (unload_dialog id))
  (vl-file-delete dcl)
  (princ)
)


;;; ---------------------------------------------------------------------------
;;; Step 6 - Clean
;;; ---------------------------------------------------------------------------

(defun yyc:text-p (o) (member (vla-get-ObjectName o) '("AcDbText" "AcDbMText")))
(defun yyc:empty-text-p (o / s)
  (and (yyc:text-p o) (setq s (yyc:try 'vla-get-TextString (list o))) (= (yyc:trim s) ""))
)

(defun yyc:t-clean (doc / kill r)
  (vlax-for blk (vla-get-Blocks doc)
    (if (yyc:own-block-p blk)
      (vlax-for o blk (if (yyc:empty-text-p o) (setq kill (cons o kill))))
    )
  )
  (foreach o kill (vl-catch-all-apply 'vla-Delete (list o)))
  (yyc:msg (strcat "Deleted " (itoa (length kill)) " empty text object(s)."))
  (repeat 3 (vl-catch-all-apply 'vla-PurgeAll (list doc)))
  (yyc:msg "Purged all unused definitions (3 passes, for nested items).")
  (setq r (vl-catch-all-apply 'vla-AuditInfo (list doc :vlax-true)))
  (yyc:msg (if (yyc:err-p r) (strcat "AUDIT failed: " (vl-catch-all-error-message r))
                             "AUDIT ran with fix on. To read its findings, type AUDIT once by hand."))
)
(defun c:YYCCLEAN () (yyc:run 'yyc:t-clean))

;;; ---------------------------------------------------------------------------
;;; Step 6 - "$0$" scan
;;; ---------------------------------------------------------------------------

(defun yyc:has-$0$ (s) (and s (= (type s) 'STR) (wcmatch s "*$0$*")))

(defun yyc:dict-hits (dict / res nm)
  (vl-catch-all-apply
    '(lambda ()
       (vlax-for it dict
         (setq nm (yyc:try 'vla-GetName (list dict it)))
         (if (yyc:has-$0$ nm) (setq res (cons nm res)))
         (if (= (vla-get-ObjectName it) "AcDbDictionary") (setq res (append res (yyc:dict-hits it))))
       )))
  res
)

;; list of ("Collection" names...) with hits only
(defun yyc:find-$0$ (doc / res hits nm)
  (foreach pr (list (cons "Layers" (vla-get-Layers doc)) (cons "Linetypes" (vla-get-Linetypes doc))
                    (cons "Text styles" (vla-get-TextStyles doc)) (cons "Dim styles" (vla-get-DimStyles doc))
                    (cons "Blocks" (vla-get-Blocks doc)) (cons "Views" (vla-get-Views doc))
                    (cons "UCS" (vla-get-UserCoordinateSystems doc)) (cons "Viewports" (vla-get-Viewports doc))
                    (cons "RegApps" (vla-get-RegisteredApplications doc)))
    (setq hits nil)
    (vlax-for x (cdr pr)
      (setq nm (yyc:try 'vla-get-Name (list x)))
      (if (and (yyc:has-$0$ nm) (not (member nm hits))) (setq hits (cons nm hits)))
    )
    (if hits (setq res (cons (cons (car pr) (reverse hits)) res)))
  )
  (if (setq hits (yyc:dict-hits (vla-get-Dictionaries doc))) (setq res (cons (cons "Dictionaries" hits) res)))
  (reverse res)
)

(defun yyc:t-find0 (doc / hits)
  (if (setq hits (yyc:find-$0$ doc))
    (progn
      (yyc:msg "WARNING - names containing \"$0$\" (leftovers from binding xrefs):")
      (foreach h hits (yyc:msg (strcat "  " (car h) ": " (yyc:join (cdr h) ", "))))
      (yyc:msg "Fix: RENAME to the standard name, or merge with YYCLAYMAP.")
    )
    (yyc:msg "PASS - no names containing \"$0$\".")
  )
)
(defun c:YYCFIND0 () (yyc:run 'yyc:t-find0))

;;; ---------------------------------------------------------------------------
;;; Step 7 - Viewports
;;; The first viewport in a layout's block is the paper-space viewport itself;
;;; every one after it is a real viewport.
;;; ---------------------------------------------------------------------------

;; Real viewports only. The sheet's own paper-space viewport is skipped: it has
;; ID 1, or it simply looks at the paper itself (its view centre = its frame centre). If an object can't be read that way, the old rule
;; (first viewport in the layout) is used as a fallback.
(defun yyc:paper-vp-p (o / e d c v)
  (setq e (yyc:try 'vlax-vla-object->ename (list o)))
  (if (and e (setq d (entget e)))
    (or (= (cdr (assoc 69 d)) 1)
        (and (setq c (cdr (assoc 10 d)) v (cdr (assoc 12 d)))
             (< (abs (- (car c) (car v))) 1e-4) (< (abs (- (cadr c) (cadr v))) 1e-4)))
    'unknown)
)

(defun yyc:vp-list (doc / res first k)
  (foreach lay (yyc:paper-layouts doc)
    (setq first T)
    (vlax-for o (vla-get-Block lay)
      (if (= (vla-get-ObjectName o) "AcDbViewport")
        (progn
          (setq k (yyc:paper-vp-p o))
          (if (if (eq k 'unknown) (not first) (not k)) (setq res (cons (cons (vla-get-Name lay) o) res)))
          (setq first nil)))
    )
  )
  (reverse res)
)

(defun yyc:scale-text (vp / s)
  (setq s (yyc:try 'vla-get-CustomScale (list vp)))
  (cond ((or (not s) (<= s 0)) "?")
        ((>= s 1.0) (strcat (rtos s 2 2) ":1"))
        (T (strcat "1:" (rtos (/ 1.0 s) 2 2))))
)

(defun yyc:delete-empty-layout2 (doc / lay)
  (if (setq lay (yyc:try 'vla-Item (list (vla-get-Layouts doc) "Layout2")))
    (if (<= (vla-get-Count (vla-get-Block lay)) 1)
      (progn
        (if (= (strcase (vla-get-Name (vla-get-ActiveLayout doc))) "LAYOUT2")
          (vla-put-ActiveLayout doc (vla-Item (vla-get-Layouts doc) "Model")))
        (yyc:msg (if (yyc:err-p (vl-catch-all-apply 'vla-Delete (list lay))) "Could not delete empty Layout2." "Deleted empty Layout2."))
      )
      (yyc:msg "Layout2 has content - left alone.")
    )
  )
)

(defun yyc:t-vplayers (doc / layers n name lo vp extra)
  (setq layers (vla-get-Layers doc) n 0)
  (yyc:delete-empty-layout2 doc)
  (foreach pr (yyc:vp-list doc)
    (setq n (1+ n) name (strcat "VIEWPORT" (itoa n)) vp (cdr pr))
    (if (not (setq lo (yyc:layer-ci layers name))) (setq lo (vla-Add layers name)))
    (vla-put-Plottable lo :vlax-false)
    (yyc:msg (if (yyc:err-p (vl-catch-all-apply 'vla-put-Layer (list vp name)))
               (strcat "  FAILED: viewport in " (car pr) " (its layer may be locked)")
               (strcat "  " (car pr) ": viewport -> " name " (no plot), scale " (yyc:scale-text vp))))
  )
  (yyc:msg (strcat (itoa n) " viewport(s) processed."))
  ;; VIEWPORT# layers left over from viewports that were deleted (e.g. after YYCSCHEDULES)
  (setq extra nil)
  (vlax-for L layers
    (if (and (wcmatch (strcase (vla-get-Name L)) "VIEWPORT#*")
             (> (atoi (substr (vla-get-Name L) 9)) n))
      (setq extra (cons (vla-get-Name L) extra))))
  (if extra
    (yyc:msg (strcat "Now empty: " (yyc:join (reverse extra) ", ") " - run YYCCLEAN to purge them.")))
)
(defun c:YYCVPLAYERS () (yyc:run 'yyc:t-vplayers))

;; --- standard scales (per profile, e.g. "1:50,1:100,1:200") ---
(defun yyc:std-scales ( / l)
  (setq l (yyc:split (yyc:get "Scales" "1:1,1:2,1:5,1:10,1:20,1:25,1:50,1:75,1:100,1:125,1:200,1:250,1:500,1:1000") ","))
  (mapcar '(lambda (x) (atof (cadr (yyc:split (yyc:trim x) ":")))) l)
)
;; nil when the viewport is at a standard scale, otherwise the nearest standard denominator
(defun yyc:scale-off (vp / s den best)
  (setq s (yyc:try 'vla-get-CustomScale (list vp)))
  (if (and s (> s 0))
    (progn
      (setq den (/ 1.0 s))
      (foreach d (yyc:std-scales)
        (if (or (not best) (< (abs (- d den)) (abs (- best den)))) (setq best d)))
      (if (and best (> (abs (- best den)) (* 0.0001 best))) best nil))
  )
)

(defun yyc:t-vplock (doc / n vp near)
  (setq n 0)
  (foreach pr (yyc:vp-list doc)
    (setq vp (cdr pr) near (yyc:scale-off vp))
    (if near
      (progn
        (yyc:msg (strcat "  " (car pr) ": scale " (yyc:scale-text vp) " is NOT a standard scale (nearest 1:" (rtos near 2 0) ")."))
        (if (and (not (yyc:batch-p)) (yyc:yes (strcat "  Set it to 1:" (rtos near 2 0) " before locking?") "Yes"))
          (progn (vl-catch-all-apply 'vla-put-DisplayLocked (list vp :vlax-false))
                 (vl-catch-all-apply 'vla-put-CustomScale (list vp (/ 1.0 near)))))))
    (if (not (yyc:err-p (vl-catch-all-apply 'vla-put-DisplayLocked (list vp :vlax-true)))) (setq n (1+ n)))
    (yyc:msg (strcat "  " (car pr) ": " (vla-get-Layer vp) ", scale " (yyc:scale-text vp) ", locked"))
  )
  (yyc:msg (strcat (itoa n) " viewport(s) locked at their current scale - check the list."))
  n
)
(defun c:YYCVPLOCK () (yyc:run 'yyc:t-vplock))

;;; ---------------------------------------------------------------------------
;;; Step 8 - Layers
;;; ---------------------------------------------------------------------------

;; make sure linetype lt exists in doc: CAA.LIN (profile, then the support path), then
;; AutoCAD's acadiso.lin; ActiveX first, the -LINETYPE command as a fallback (active drawing).
(defun yyc:lt-have (lts lt) (yyc:try 'vla-Item (list lts lt)))
(defun yyc:load-lt (doc lt / lts files)
  (setq lts (vla-get-Linetypes doc))
  (if (not (yyc:lt-have lts lt))
    (progn
      (setq files (vl-remove nil (list (yyc:get "LinFile" nil) (findfile "CAA.LIN") (findfile "YYC Approved Linetypes.lin") (findfile "acadiso.lin"))))
      (foreach f files (if (not (yyc:lt-have lts lt)) (vl-catch-all-apply 'vla-Load (list lts lt f))))
      (if (and (not (yyc:lt-have lts lt)) (yyc:active-p doc))
        (foreach f files
          (if (not (yyc:lt-have lts lt))
            (progn (setvar "CMDECHO" 0)
                   (vl-catch-all-apply 'command (list "._-LINETYPE" "_Load" lt f ""))
                   (while (> (getvar "CMDACTIVE") 0) (command ""))))))))
  (if (yyc:lt-have lts lt) T nil)
)

;; --- copy colour / linetype / lineweight from the YYC Layer Reference ---
;; give one layer the colour, linetype, Default lineweight, plot and description of the
;; YYC standard. Modifier layers (A-WALL-DEMO) only get lineweight and plot. Returns
;; nil (nothing to do / not YYC), T (fixed), or a string (fixed, but linetype missing).
(setq *yyc-mod-names*
  '(("DEMO" . "Demolition") ("EXST" . "Existing to Remain") ("FUTR" . "Future Work") ("MOVE" . "Items to be Moved")
    ("NEWW" . "New Work") ("NICN" . "Not in Contract") ("NPLT" . "No Plot") ("PRPS" . "Proposed Work")
    ("RELO" . "Items to be Relocated") ("TEMP" . "Temporary Work") ("ABDN" . "Abandoned") ("ELEV" . "Elevation")
    ("EQPM" . "Equipment") ("IDEN" . "Identification Tags") ("PATT" . "Patterns") ("RMVD" . "Removed")
    ("SYMB" . "Symbols") ("TEXT" . "Text")))
(setq *yyc-desc-n* 0 *yyc-lt-lost* nil)
(defun yyc:report-lt-lost ()
  (if *yyc-lt-lost*
    (progn
      (yyc:msg (strcat "NOTE - these layers had a dashed / special linetype from Revit and are now Continuous (the YYC standard): "
                       (yyc:join (reverse *yyc-lt-lost*) ", ")))
      (yyc:msg "  If that dashing meant something (callout boundaries, match lines), move those objects to a layer whose")
      (yyc:msg "  YYC linetype is dashed (match lines: A-ANNO-CONS-MTCH), or ask YYC to approve one and add it to YYC Approved Layers.csv.")))
  (setq *yyc-lt-lost* nil)
)

(defun yyc:sync-layer (doc L std / u ref pos modp lt res d)
  (setq u (strcase (vla-get-Name L)) ref (assoc u std))
  (if (and (not ref) (= (yyc:layer-status u std) "MODIFIER") (setq pos (vl-string-position 45 u nil T)))
    (setq ref (assoc (substr u 1 pos) std) modp T))
  (if (and ref (/= (yyc:layer-diff L std) ""))
    (progn
      (setq res T lt (nth 2 ref))
      (if (not modp)
        (progn
          (if (and (/= (strcase (vla-get-Linetype L)) (strcase lt)) (= (strcase lt) "CONTINUOUS")
                   (not (member (strcase (vla-get-Linetype L)) '("BYLAYER" "BYBLOCK"))))
            (setq *yyc-lt-lost* (cons (strcat (vla-get-Name L) " (" (vla-get-Linetype L) ")") *yyc-lt-lost*)))
          (vl-catch-all-apply 'vla-put-Color (list L (nth 1 ref)))
          (yyc:load-lt doc lt)
          (if (yyc:err-p (vl-catch-all-apply 'vla-put-Linetype (list L lt)))
            (setq res (strcat (vla-get-Name L) " (linetype " lt " not loaded)")))))
      (vl-catch-all-apply 'vla-put-Lineweight (list L (nth 3 ref)))
      (cond ((wcmatch u "*-NPLT") (vla-put-Plottable L :vlax-false))
            ((nth 4 ref) (vl-catch-all-apply 'vla-put-Plottable (list L (nth 4 ref)))))))
  ;; description: fill it in when the layer has none (A-WALL-DEMO gets "Wall - Demolition")
  (if (and ref (nth 5 ref) (/= (nth 5 ref) "")
           (= (yyc:trim (cond ((yyc:try 'vla-get-Description (list L))) (""))) ""))
    (progn
      (setq d (if modp (strcat (nth 5 ref) " - " (cond ((cdr (assoc (substr u (+ pos 2)) *yyc-mod-names*))) ((substr u (+ pos 2))))) (nth 5 ref)))
      (if (not (yyc:err-p (vl-catch-all-apply 'vla-put-Description (list L d))))
        (progn (setq *yyc-desc-n* (1+ (cond (*yyc-desc-n*) (0)))) (if (not res) (setq res T))))))
  res
)

(defun yyc:t-laysync (doc / std n r fail)
  (setq std (yyc:std-layers) n 0 *yyc-desc-n* 0 *yyc-lt-lost* nil)
  (if (not std)
    (yyc:msg "No YYC standard set for this profile - open YYCPROFILES and set Standard layers CSV.")
    (progn
      (vlax-for L (vla-get-Layers doc)
        (if (setq r (yyc:sync-layer doc L std))
          (progn (setq n (1+ n)) (if (= (type r) 'STR) (setq fail (cons r fail))))))
      (yyc:msg (strcat (itoa n) " layer(s) set to the colour, linetype, lineweight and plot setting of the YYC standard."))
      (if (> *yyc-desc-n* 0) (yyc:msg (strcat "  " (itoa *yyc-desc-n*) " layer description(s) filled in from the standard.")))
      (yyc:report-lt-lost)
      (if fail (yyc:msg (strcat "WARNING - linetype missing: " (yyc:join (reverse fail) ", ") " - check Linetype file (CAA.LIN) in YYCPROFILES.")))
    )
  )
)
(defun c:YYCLAYSYNC () (yyc:run 'yyc:t-laysync))
;; --- YYCMAKEREF: build the YYC Layer Reference drawing (and .dws) from the standard CSV.
;; Run once in a NEW, empty drawing. Every layer from the CADD Manual v6.2 with its
;; colour, linetype, Default lineweight, plot setting and description. Saved into the kit
;; folder as the LAYTRANS target, so the DWG and the tools always agree.
(defun yyc:t-makeref (doc / std csv kit layers lin L n miss path dws r)
  (setq csv (yyc:get "StdCSV" nil) kit (yyc:get "Kit" nil) lin (yyc:get "LinFile" nil) n 0)
  (cond
    ((not (and csv (findfile csv))) (yyc:msg "Set the Standard layers CSV in YYCPROFILES first."))
    ((not (and kit (vl-file-directory-p kit))) (yyc:msg "Set the Kit folder in YYCPROFILES first."))
    ((> (vla-get-Count (vla-get-ModelSpace doc)) 0)
     (yyc:msg "This drawing isn't empty. Start a new drawing (QNEW, acadiso.dwt), then run YYCMAKEREF."))
    (T
     (setq std (append (yyc:std-from-csv csv) (yyc:std-approved csv)) layers (vla-get-Layers doc))
     (foreach r (vl-sort std '(lambda (a b) (< (car a) (car b))))
       (setq L (if (yyc:layer-ci layers (car r)) (yyc:layer-ci layers (car r)) (vla-Add layers (car r))))
       (vl-catch-all-apply 'vla-put-Color (list L (nth 1 r)))
       (yyc:load-lt doc (nth 2 r))
       (if (yyc:err-p (vl-catch-all-apply 'vla-put-Linetype (list L (nth 2 r))))
         (setq miss (cons (strcat (car r) " (" (nth 2 r) ")") miss)))
       (vl-catch-all-apply 'vla-put-Lineweight (list L acLnWtByLwDefault))
       (if (nth 4 r) (vl-catch-all-apply 'vla-put-Plottable (list L (nth 4 r))))
       (if (and (nth 5 r) (/= (nth 5 r) "")) (vl-catch-all-apply 'vla-put-Description (list L (nth 5 r))))
       (setq n (1+ n)))
     (vla-put-ActiveLayer doc (vla-Item layers "0"))
     (setq path (strcat kit "\\YYC Layer Reference - CADD Manual v6.2.dwg")
           dws  (strcat kit "\\YYC Layer Reference - CADD Manual v6.2.dws"))
     (setq r (vl-catch-all-apply 'vla-SaveAs (list doc path (cond ((eval 'ac2018_dwg)) (64)))))
     (if (yyc:err-p r) (setq r (vl-catch-all-apply 'vla-SaveAs (list doc path))))
     (cond
       ((yyc:err-p r) (yyc:msg (strcat "Could not save " path ": " (vl-catch-all-error-message r))))
       (T
        (if (findfile dws) (vl-file-delete dws))
        (vl-file-copy path dws)
        (yyc:set "LayerRef" path)
        (yyc:msg (strcat (itoa n) " YYC layers written to " path))
        (yyc:msg "  ...and the same file as a .dws standards file next to it. Load either one in LAYTRANS.")
        (if miss (yyc:msg (strcat "  WARNING - linetype could not be loaded for " (itoa (length miss)) " layer(s), left Continuous: " (yyc:join (reverse miss) ", ") ". Check Linetype file (CAA.LIN) in YYCPROFILES and that CAA.SHX is on the support path.")))
        (yyc:msg (strcat "  Linetype file used: " (cond ((yyc:get "LinFile" nil)) ((findfile "CAA.LIN")) ("NONE SET")))))))
  )
)
(defun c:YYCMAKEREF () (yyc:run 'yyc:t-makeref))

(defun yyc:t-lwdefault (doc / n)
  (setq n 0)
  (vlax-for L (vla-get-Layers doc)
    (if (and (not (vl-string-search "|" (vla-get-Name L))) (/= (vla-get-Lineweight L) acLnWtByLwDefault))
      (if (not (yyc:err-p (vl-catch-all-apply 'vla-put-Lineweight (list L acLnWtByLwDefault)))) (setq n (1+ n)))
    )
  )
  (yyc:msg (strcat (itoa n) " layer(s) set to lineweight Default. Object-level lineweights: SETBYLAYER."))
)
(defun c:YYCLWDEFAULT () (yyc:run 'yyc:t-lwdefault))

;; YYC standard as ((NAME colour LINETYPE lineweight plot) ...). Read from the
;; standard layer CSV (taken from the Layer Comparison workbook) when the profile has
;; one, otherwise from the YYC Layer Reference drawing via ObjectDBX (no plot info).
(defun yyc:unq (x) (yyc:trim (vl-string-trim "\"" (yyc:trim x))))

(defun yyc:lw-parse (x / u)
  (setq u (strcase (yyc:unq x)))
  (cond ((wcmatch u "*DEFAULT*") -3) ((wcmatch u "BYLAYER") -1) ((wcmatch u "BYBLOCK") -2)
        ((wcmatch u "LINEWEIGHT*") (atoi (substr u 11))) (T (atoi u)))
)

(defun yyc:std-from-csv (path / f line c res)
  (if (setq f (open path "r"))
    (progn
      (while (setq line (read-line f))
        (setq c (mapcar 'yyc:unq (yyc:split line ",")))
        (if (and (>= (length c) 5) (/= (strcase (car c)) "LAYER") (/= (car c) ""))
          (setq res (cons (list (strcase (car c)) (if (= (strcase (nth 1 c)) "WHITE") 7 (atoi (nth 1 c))) (strcase (nth 2 c)) (yyc:lw-parse (nth 3 c))
                                (if (= (strcase (nth 4 c)) "FALSE") :vlax-false :vlax-true)
                                (vl-string-trim "\" " (yyc:join (reverse (cdr (reverse (cdr (cddddr c))))) ","))) res))))
      (close f)))
  res
)

;; layers YYC approved outside the manual: "*Approved*Layers*.csv" next to the standard CSV
(defun yyc:std-approved (csv / dir res)
  (setq dir (vl-filename-directory csv))
  (foreach f (vl-directory-files dir "*Approved*Layers*.csv" 1)
    (setq res (append res (yyc:std-from-csv (strcat dir "\\" f)))))
  res
)

(defun yyc:std-layers ( / csv path dbx res key)
  (setq csv (yyc:get "StdCSV" nil) path (yyc:get "LayerRef" nil))
  (setq key (if (and csv (findfile csv)) csv path))
  (cond
    ((and *yyc-std* (= *yyc-std-path* key)) *yyc-std*)
    ((and csv (findfile csv))
     (setq res (append (yyc:std-from-csv csv) (yyc:std-approved csv)) *yyc-std* res *yyc-std-path* key) res)
    ((not (and path (findfile path))) nil)
    ((setq dbx (yyc:dbx-open path))
     (vlax-for L (vla-get-Layers dbx)
       (setq res (cons (list (strcase (vla-get-Name L)) (vla-get-Color L) (strcase (vla-get-Linetype L)) (vla-get-Lineweight L) nil) res)))
     (yyc:dbx-close dbx)
     (setq *yyc-std* res *yyc-std-path* key)
     res)
  )
)

;; "SYSTEM" / "STANDARD" / "MODIFIER" / "NON-STANDARD" / "XREF"
(defun yyc:layer-status (name std / u pos)
  (setq u (strcase name))
  (cond
    ((vl-string-search "|" u) "XREF")
    ((member u '("0" "DEFPOINTS")) "SYSTEM")
    ((assoc u std) "STANDARD")
    ((and (setq pos (vl-string-position 45 u nil T))
          (member (substr u (+ pos 2)) *yyc-modifiers*)
          (assoc (substr u 1 pos) std)) "MODIFIER")
    (T "NON-STANDARD")
  )
)

(defun yyc:lw-text (lw)
  (cond ((= lw -3) "Default") ((= lw -2) "ByBlock") ((= lw -1) "ByLayer") (T (strcat (rtos (/ lw 100.0) 2 2) "mm")))
)

;; A layer with a modifier (A-WALL-DEMO) is checked against its parent (A-WALL) for
;; lineweight and plot only. The manual gives no colour or linetype for modifiers, so a
;; red dashed DEMO layer is left as you set it. -NPLT layers must be no-plot.
(defun yyc:layer-diff (L std / u pos ref out modp plot)
  (setq u (strcase (vla-get-Name L)) ref (assoc u std))
  (if (and (not ref) (setq pos (vl-string-position 45 u nil T))) (setq ref (assoc (substr u 1 pos) std) modp T))
  (if ref
    (progn
      (if (and (not modp) (/= (vla-get-Color L) (nth 1 ref)))
        (setq out (cons (strcat "colour " (itoa (vla-get-Color L)) " should be " (itoa (nth 1 ref))) out)))
      (if (and (not modp) (/= (strcase (vla-get-Linetype L)) (nth 2 ref)))
        (setq out (cons (strcat "linetype " (vla-get-Linetype L) " should be " (nth 2 ref)) out)))
      (if (/= (vla-get-Lineweight L) (nth 3 ref))
        (setq out (cons (strcat "lineweight " (yyc:lw-text (vla-get-Lineweight L)) " should be " (yyc:lw-text (nth 3 ref))) out)))
      (setq plot (if (wcmatch u "*-NPLT") :vlax-false (nth 4 ref)))
      (if (and plot (/= (vla-get-Plottable L) plot))
        (setq out (cons (if (= plot :vlax-true) "should plot" "should be no-plot") out)))
    )
  )
  (yyc:join (reverse out) "; ")
)

(defun yyc:q (s) (strcat "\"" (vl-string-translate "\"" "'" s) "\""))
(defun yyc:yn (v) (if (= v :vlax-true) "Y" "N"))

;; rewrite a shared CSV, replacing this drawing's rows
(defun yyc:csv-replace (csv me header newrows / f old lines)
  (if (setq f (open csv "r"))
    (progn
      (while (setq old (read-line f))
        (if (/= (substr old 1 (+ 2 (strlen me))) (yyc:q me)) (setq lines (cons old lines))))
      (close f)
      (setq lines (reverse lines))
    )
  )
  (if (not lines) (setq lines (list header)))
  (if (setq f (open csv "w"))
    (progn (foreach x (append lines newrows) (write-line x f)) (close f) T)
    nil)
)

(defun yyc:t-layexport (doc / std cnt me rows nm csv)
  (setq std (yyc:std-layers) me (yyc:dname doc) cnt (nth 0 (yyc:walk doc)))
  (if (not std) (yyc:msg "No layer reference set (YYCSETUP) - exporting without the standard check."))
  (vlax-for L (vla-get-Layers doc)
    (setq nm (vla-get-Name L))
    (setq rows (cons (yyc:join (list (yyc:q me) (yyc:q nm)
                   (itoa (cond ((cdr (assoc (strcase nm) cnt))) (0)))
                   (if std (yyc:layer-status nm std) "?")
                   (yyc:q (if std (yyc:layer-diff L std) ""))
                   (yyc:yn (vla-get-LayerOn L)) (yyc:yn (vla-get-Freeze L)) (yyc:yn (vla-get-Lock L)) (yyc:yn (vla-get-Plottable L))
                   (itoa (vla-get-Color L)) (yyc:q (vla-get-Linetype L)) (yyc:lw-text (vla-get-Lineweight L))) ",") rows))
  )
  (setq csv (strcat (yyc:dfolder doc) "YYC_LayerExport.csv"))
  (if (yyc:csv-replace csv me "Drawing,Layer,Objects,Standard,Differences,On,Frozen,Locked,Plot,Colour,Linetype,Lineweight" (reverse rows))
    (yyc:msg (strcat "Layers written to " csv " - filter Standard = NON-STANDARD, or Objects = 0 for blank layers."))
    (yyc:msg (strcat "ERROR: could not write " csv " (open in Excel?)"))
  )
)
(defun c:YYCLAYEXPORT () (yyc:run 'yyc:t-layexport))

;; --- write / extend the layer map CSV from this drawing ---
;; Every layer that isn't in the YYC standard gets a row with a suggested target:
;;  1. Revit's number suffix removed (A-WALL-6 -> A-WALL) if that is a YYC layer
;;  2. otherwise the nearest YYC parent layer (A-FLOR-STRS-1 -> A-FLOR-STRS or A-FLOR)
;;  3. otherwise blank - you fill it in. Existing rows are kept, so the map grows
;;     into a library as you go sheet by sheet.
(defun yyc:strip-num (u / pos)
  (while (and (setq pos (vl-string-position 45 u nil T))
              (> (strlen (substr u (+ pos 2))) 0)
              (vl-every '(lambda (c) (<= 48 c 57)) (vl-string->list (substr u (+ pos 2)))))
    (setq u (substr u 1 pos)))
  u
)

(defun yyc:suggest (name std / u a b pos)
  (setq u (strcase name) a (yyc:strip-num u))
  (cond
    ((and (/= a u) (member (yyc:layer-status a std) '("STANDARD" "MODIFIER"))) (list a "number suffix removed"))
    (T
     (setq b a)
     (while (and (setq pos (vl-string-position 45 b nil T)) (not (assoc b std)))
       (setq b (substr b 1 pos)))
     (if (and (assoc b std) (/= b u)) (list b "nearest YYC parent layer - check") (list "" "no match - fill in")))
  )
)

(defun yyc:t-laymapmake (doc / std csv old rows cnt f nm u key st sug added auto have tot)
  (setq std (yyc:std-layers) cnt (nth 0 (yyc:walk doc)) added 0 auto 0)
  (if (not std)
    (yyc:msg "No YYC standard set for this profile (standard CSV or Layer Reference) - open YYCPROFILES and set it.")
    (progn
      (setq csv (cond ((yyc:get "LayerMap" nil)) (T (strcat (yyc:get "Kit" (vla-get-Path doc)) "\\YYC_LayerMap.csv"))))
      (if (setq f (open csv "r"))
        (progn (while (setq nm (read-line f)) (setq old (cons nm old))) (close f) (setq old (reverse old))))
      (if (not old) (setq old (list "old,new,how,objects,first seen in")))
      (setq rows old have (mapcar '(lambda (r) (strcase (yyc:unq (car (yyc:split r ","))))) old))
      ;; one row per base name (Revit numbers stripped); numbered variants follow it
      (vlax-for L (vla-get-Layers doc)
        (setq nm (vla-get-Name L) u (strcase nm) key (yyc:strip-num u) st (yyc:layer-status u std))
        (cond
          ((/= st "NON-STANDARD"))
          ((wcmatch u "VIEWPORT#*,T-TTLB*"))
          ((and (/= key u) (member (yyc:layer-status key std) '("STANDARD" "MODIFIER"))) (setq auto (1+ auto)))
          ((or (member u have) (member key have)))
          (T
           (setq sug (yyc:suggest key std) added (1+ added) have (cons key have))
           ;; objects on the base layer and all its numbered variants
           (setq tot 0)
           (foreach c cnt (if (= (yyc:strip-num (car c)) key) (setq tot (+ tot (cdr c)))))
           (setq rows (append rows (list (yyc:join (list key (car sug) (cadr sug) (itoa tot) (yyc:dname doc)) ",")))))
        )
      )
      (if (setq f (open csv "w"))
        (progn
          (foreach r rows (write-line r f)) (close f)
          (yyc:set "LayerMap" csv)
          (yyc:msg (strcat (itoa auto) " layer(s) map automatically (YYC name + Revit number, e.g. A-WALL-6 -> A-WALL) - nothing to do."))
          (yyc:msg (strcat (itoa added) " base name(s) added to " csv " for you to decide."))
          (yyc:msg "Open it in Excel: check rows marked 'check', fill blank 'new' cells with a YYC layer (your Layer Comparison")
          (yyc:msg "workbook's YYCStandard tab lists them), save as CSV. Each base name covers its -1, -2 ... variants too.")
          (yyc:msg "Then YYCLAYMAP on each sheet (or Mapcsv in YYCBATCH), and YYCLAYSYNC for colours/linetypes."))
        (yyc:msg (strcat "ERROR: could not write " csv " - is it open in Excel?")))
    )
  )
)
(defun c:YYCLAYMAPMAKE () (yyc:run 'yyc:t-laymapmake))

;;; ---------------------------------------------------------------------------
;;; Step 8 - Layer Mapper workbook (Excel, through ActiveX)
;;; YYCLAYXL   : sends this drawing's layers to the workbook's Map sheet, each with a
;;;              status and a suggested YYC layer (remembered ones first)
;;; YYCLAYMAP  : reads your choices back, merges the layers, remembers them in Library
;;; ---------------------------------------------------------------------------

(defun yyc:xv (v) (if (= (type v) 'variant) (vlax-variant-value v) v))
(defun yyc:xs (v / x)   ; cell value as a trimmed string ("" when empty)
  (setq x (yyc:xv v))
  (cond ((null x) "") ((= (type x) 'STR) (yyc:trim x)) ((numberp x) (rtos x 2 0)) (T ""))
)

;; open (or reuse) the workbook in Excel; returns (excel workbook) or nil
(defun yyc:xl-open (path visible / xl wbs wb i w)
  (setq xl (vl-catch-all-apply 'vlax-get-or-create-object (list "Excel.Application")))
  (if (or (yyc:err-p xl) (not xl))
    (progn (yyc:msg "Excel could not be started on this machine.") nil)
    (progn
      (setq wbs (vlax-get-property xl 'Workbooks) i 1)
      (repeat (vlax-get-property wbs 'Count)
        (setq w (vlax-get-property wbs 'Item i) i (1+ i))
        (if (= (strcase (vlax-get-property w 'FullName)) (strcase path)) (setq wb w)))
      (if (not wb)
        (if (yyc:err-p (setq wb (vl-catch-all-apply 'vlax-invoke-method (list wbs 'Open path))))
          (progn (yyc:msg (strcat "Could not open " path)) (setq wb nil))))
      (if visible (vlax-put-property xl 'Visible :vlax-true))
      (if wb (list xl wb))
    )
  )
)
(defun yyc:xl-sheet (wb name) (yyc:try 'vlax-get-property (list (vlax-get-property wb 'Worksheets) 'Item name)))
(defun yyc:xl-put (ws addr val) (vlax-put-property (vlax-get-property ws 'Range addr) 'Value2 val))
(defun yyc:xl-rows (ws addr / v)   ; a block of cells as a list of row lists of strings
  (setq v (yyc:xv (vlax-get-property (vlax-get-property ws 'Range addr) 'Value2)))
  (if (= (type v) 'SAFEARRAY) (mapcar '(lambda (r) (mapcar 'yyc:xs r)) (vlax-safearray->list v)))
)

;; Library sheet: ((EXPORT-BASE . YYC layer) ...), row numbers kept for updating
(defun yyc:xl-library (wb / ws rows i res)
  (if (setq ws (yyc:xl-sheet wb "Library"))
    (progn
      (setq rows (yyc:xl-rows ws "A2:B5000") i 2)
      (foreach r rows
        (if (and (/= (car r) "") (/= (cadr r) "")) (setq res (cons (list (strcase (car r)) (cadr r) i) res)))
        (setq i (1+ i)))))
  res
)

(defun yyc:t-layxl (doc / book std cnt x wb ws lib rows u key st sug row n)
  (setq std (yyc:std-layers) cnt (nth 0 (yyc:walk doc)))
  (cond
    ((not std) (yyc:msg "No YYC standard in this profile - YYCPROFILES, set Standard layers CSV."))
    ((not (setq book (yyc:need-file "LayerBook" "Select the YYC Layer Mapper workbook" "xlsx"))) nil)
    ((not (setq x (yyc:xl-open book T))) nil)
    ((not (setq ws (yyc:xl-sheet (setq wb (cadr x)) "Map"))) (yyc:msg "The workbook has no Map sheet - is it the YYC Layer Mapper?"))
    (T
     (setq lib (yyc:xl-library wb))
     ;; one row per layer: export name, objects, status, suggestion, choice (prefilled)
     (vlax-for L (vla-get-Layers doc)
       (setq u (strcase (vla-get-Name L)) key (yyc:strip-num u) st (yyc:layer-status u std))
       (if (not (or (= st "XREF") (wcmatch u "VIEWPORT#*")))
         (progn
           (setq sug
             (cond
               ((member st '("STANDARD" "MODIFIER" "SYSTEM")) (list (if (= st "MODIFIER") "OK - modifier" "OK") (vla-get-Name L)))
               ((assoc u lib) (list "REMEMBERED" (cadr (assoc u lib))))
               ((assoc key lib) (list "REMEMBERED" (cadr (assoc key lib))))
               ((and (/= key u) (member (yyc:layer-status key std) '("STANDARD" "MODIFIER"))) (list "AUTO - Revit number" key))
               (T (list "CHOOSE" (car (yyc:suggest key std))))))
           (setq rows (cons (list (vla-get-Name L) (cond ((cdr (assoc u cnt))) (0)) (car sug) (cadr sug)
                                  (if (member st '("STANDARD" "MODIFIER")) (yyc:layer-diff L std) "")) rows)))))
     ;; CHOOSE first (most objects first), then REMEMBERED, AUTO, OK
     (setq rows (vl-sort rows '(lambda (a b / ra rb)
                   (setq ra (vl-position (substr (caddr a) 1 4) '("CHOO" "REME" "AUTO" "OK -" "OK"))
                         rb (vl-position (substr (caddr b) 1 4) '("CHOO" "REME" "AUTO" "OK -" "OK")))
                   (if (= ra rb) (> (cadr a) (cadr b)) (< ra rb)))))
     (vlax-invoke-method (vlax-get-property ws 'Range "A2:E5000") 'ClearContents)
     (vlax-invoke-method (vlax-get-property ws 'Range "H2:I5000") 'ClearContents)
     (setq n 1)
     (foreach r rows
       (setq n (1+ n) row (itoa n))
       (yyc:xl-put ws (strcat "A" row) (car r))
       (yyc:xl-put ws (strcat "B" row) (cadr r))
       (yyc:xl-put ws (strcat "C" row) (caddr r))
       (yyc:xl-put ws (strcat "D" row) (cadddr r))
       (yyc:xl-put ws (strcat "E" row) (cadddr r))
       (yyc:xl-put ws (strcat "H" row) (nth 4 r)))
     (yyc:xl-put ws "J1" (yyc:dname doc))
     (vlax-invoke-method ws 'Activate)
     (vl-catch-all-apply 'vlax-invoke-method (list wb 'Save))
     (yyc:msg (strcat (itoa (length rows)) " layers sent to the Map sheet of " book))
     (yyc:msg (strcat "  To choose: " (itoa (length (vl-remove-if-not '(lambda (r) (= (caddr r) "CHOOSE")) rows)))
                      "   remembered: " (itoa (length (vl-remove-if-not '(lambda (r) (= (caddr r) "REMEMBERED")) rows)))
                      "   automatic: " (itoa (length (vl-remove-if-not '(lambda (r) (wcmatch (caddr r) "AUTO*")) rows)))))
     (yyc:msg "In Excel: for each CHOOSE row pick a YYC layer in the yellow column (or leave it blank to walk it by hand), save,")
     (yyc:msg "then come back and run YYCLAYMAP.")
     (if (vl-some '(lambda (r) (/= (nth 4 r) "")) rows)
       (yyc:msg (strcat "  " (itoa (length (vl-remove-if '(lambda (r) (= (nth 4 r) "")) rows)))
                        " YYC-named layer(s) have the wrong colour / linetype / lineweight (column H) - YYCLAYSYNC fixes them.")))
    )
  )
)
(defun c:YYCLAYXL () (yyc:run 'yyc:t-layxl))

;; your choices from the workbook: ((LAYER . target) ...); also updates the Library
(defun yyc:xl-choices (doc remember / book x wb ws rows lib libws res next k i)
  (if (and (setq book (yyc:get "LayerBook" nil)) (findfile book) (setq x (yyc:xl-open book nil)))
    (progn
      (setq wb (cadr x) ws (yyc:xl-sheet wb "Map") lib (yyc:xl-library wb) libws (yyc:xl-sheet wb "Library"))
      (if ws
        (foreach r (yyc:xl-rows ws "A2:E5000")
          (if (and (/= (car r) "") (/= (nth 4 r) "") (/= (strcase (nth 4 r)) (strcase (car r))))
            (setq res (cons (cons (strcase (car r)) (nth 4 r)) res)))))
      ;; remember: export base name -> YYC layer, newest wins
      (if (and remember libws res)
        (progn
          (setq next 2)
          (foreach r (yyc:xl-rows libws "A2:A5000") (if (/= (car r) "") (setq next (1+ next))))
          (foreach m res
            (setq k (yyc:strip-num (car m)))
            (if (setq i (caddr (assoc k lib)))
              (yyc:xl-put libws (strcat "B" (itoa i)) (cdr m))
              (progn
                (yyc:xl-put libws (strcat "A" (itoa next)) k)
                (yyc:xl-put libws (strcat "B" (itoa next)) (cdr m))
                (setq lib (cons (list k (cdr m) next) lib) next (1+ next))))
            (yyc:xl-put libws (strcat "C" (itoa (caddr (assoc k lib)))) (yyc:stamp))
            (yyc:xl-put libws (strcat "D" (itoa (caddr (assoc k lib)))) (yyc:dname doc)))
          (vl-catch-all-apply 'vlax-invoke-method (list wb 'Save))
          (yyc:msg (strcat (itoa (length res)) " choice(s) remembered in the workbook's Library sheet."))))
      ;; Library rows also count, so other sheets get remembered mappings
      (foreach l lib (if (not (assoc (car l) res)) (setq res (append res (list (cons (car l) (cadr l)))))))
    )
  )
  res
)

;; --- CSV layer map ---
(defun yyc:read-map (path / f line parts res)
  (if (setq f (open path "r"))
    (progn
      (while (setq line (read-line f))
        (setq parts (mapcar 'yyc:unq (yyc:split line ",")))
        (if (and (>= (length parts) 2) (/= (car parts) "") (/= (cadr parts) "")
                 (/= (strcase (car parts)) "OLD") (/= (strcase (car parts)) "EXPORT LAYER"))
          (setq res (cons (cons (car parts) (cadr parts)) res)))
      )
      (close f)
    )
  )
  (reverse res)
)

(defun yyc:retag (o map / hit n)
  (setq n 0)
  (if (and (setq hit (assoc (strcase (vla-get-Layer o)) map))
           (not (yyc:err-p (vl-catch-all-apply 'vla-put-Layer (list o (cdr hit))))))
    (setq n 1))
  (if (and (= (vla-get-ObjectName o) "AcDbBlockReference") (= (vla-get-HasAttributes o) :vlax-true))
    (foreach a (vlax-invoke o 'GetAttributes) (setq n (+ n (yyc:retag a map)))))
  n
)

(defun yyc:t-laymap (doc / layers std csv pairs map oldL newL tgt u moved del kept created auto fromcsv targets L2 r2 ltfail)
  (setq layers (vla-get-Layers doc) std (yyc:std-layers) moved 0 del 0 auto 0 fromcsv 0)
  (setq pairs (yyc:xl-choices doc (not (yyc:batch-p))))
  (if pairs (yyc:msg (strcat "Using your choices from the Layer Mapper workbook (" (vl-filename-base (yyc:get "LayerBook" "")) ")."))
    (progn
      (setq csv (yyc:get "LayerMap" nil))
      (if (and csv (findfile csv)) (setq pairs (mapcar '(lambda (p) (cons (strcase (car p)) (cdr p))) (yyc:read-map csv))))))
  (if (not std) (yyc:msg "No YYC standard set - only the CSV map is used (YYCPROFILES, set Standard layers CSV)."))
  ;; decide a target for every layer: CSV row for the exact name, then CSV row for the
  ;; name without its Revit number, then the YYC layer the number was added to
  (vlax-for L layers
    (setq u (strcase (vla-get-Name L)) tgt nil)
    (cond
      ((assoc u pairs) (setq tgt (cdr (assoc u pairs)) fromcsv (1+ fromcsv)))
      ((and std (member (yyc:layer-status u std) '("STANDARD" "MODIFIER" "SYSTEM" "XREF"))) nil)
      ((assoc (yyc:strip-num u) pairs) (setq tgt (cdr (assoc (yyc:strip-num u) pairs)) fromcsv (1+ fromcsv)))
      ((and std (/= (yyc:strip-num u) u) (member (yyc:layer-status (yyc:strip-num u) std) '("STANDARD" "MODIFIER")))
       (setq tgt (yyc:strip-num u) auto (1+ auto))))
    (if (and tgt (/= (strcase tgt) u))
      (setq map (cons (cons u tgt) map)))
  )
  ;; make the targets, unlock the sources
  (setq map (mapcar '(lambda (m)
                       (if (not (setq newL (yyc:layer-ci layers (cdr m))))
                         (setq newL (vla-Add layers (cdr m)) created (cons (cdr m) created)))
                       (vla-put-Lock (vla-Item layers (car m)) :vlax-false)
                       (cons (car m) (vla-get-Name newL)))
                    map))
  (if (not map)
    (yyc:msg "Nothing to merge - every layer is already standard, or has no mapping yet (YYCLAYMAPMAKE lists them).")
    (progn
      (if (assoc (strcase (vla-get-Name (vla-get-ActiveLayer doc))) map)
        (vla-put-ActiveLayer doc (vla-Item layers "0")))
      (vlax-for blk (vla-get-Blocks doc)
        (if (yyc:own-block-p blk) (vlax-for o blk (setq moved (+ moved (yyc:retag o map))))))
      (foreach m (reverse map)
        (if (yyc:err-p (vl-catch-all-apply '(lambda () (vla-Delete (vla-Item layers (car m))))))
          (setq kept (cons (car m) kept))
          (setq del (1+ del)))
        (yyc:msg (strcat "  " (car m) " -> " (cdr m))))
      (yyc:msg (strcat "Moved " (itoa moved) " object(s); merged " (itoa del) " layer(s): "
                       (itoa auto) " by removing the Revit number, " (itoa fromcsv) " from your choices."))
      ;; the target layers get their YYC colour, linetype, lineweight and plot straight away
      (if std
        (foreach t2 (vl-remove-if '(lambda (x) (member x (cdr (member x targets)))) (setq targets (mapcar 'cdr map)))
          (if (setq L2 (yyc:layer-ci layers t2)) (if (= (type (setq r2 (yyc:sync-layer doc L2 std))) 'STR) (setq ltfail (cons r2 ltfail))))))
      (yyc:msg "Target layers set to their YYC colour, linetype, lineweight and plot.")
      (yyc:report-lt-lost)
      (if ltfail (yyc:msg (strcat "WARNING - linetype missing: " (yyc:join ltfail ", ") " - check Linetype file (CAA.LIN) in YYCPROFILES.")))
      (if kept (yyc:msg (strcat "Emptied but not deleted (YYCCLEAN purges them): " (yyc:join kept ", "))))
      (if created (yyc:msg (strcat "Created: " (yyc:join created ", "))))
    )
  )
  (yyc:msg "Layers still not YYC? YYCLAYWALK shows them one at a time so you can map each one.")
)
(defun c:YYCLAYMAP () (yyc:run 'yyc:t-laymap))

;;; ---------------------------------------------------------------------------
;;; YYCLAYWALK - the layers left over after YYCLAYMAP, one at a time.
;;; A window lists every layer that still isn't YYC. Pick one, see it on its own
;;; (Show it), pick the YYC layer it belongs on, Map. Each choice is merged right
;;; away and remembered in the Layer Mapper workbook's Library for next time.
;;; ---------------------------------------------------------------------------

;; move everything on layer FROM onto layer TO (made if missing), delete FROM
(defun yyc:merge-one (doc from to / layers src dst n)
  (setq layers (vla-get-Layers doc) n 0)
  (if (not (setq dst (yyc:layer-ci layers to))) (setq dst (vla-Add layers to)))
  (setq src (yyc:layer-ci layers from))
  (if src
    (progn
      (vla-put-Lock src :vlax-false)
      (if (= (strcase (vla-get-Name (vla-get-ActiveLayer doc))) (strcase from))
        (vla-put-ActiveLayer doc (vla-Item layers "0")))
      (vlax-for blk (vla-get-Blocks doc)
        (if (yyc:own-block-p blk)
          (vlax-for o blk (setq n (+ n (yyc:retag o (list (cons (strcase from) (vla-get-Name dst)))))))))
      (vl-catch-all-apply 'vla-Delete (list src))))
  (if *yyc-wstd* (yyc:sync-layer doc dst *yyc-wstd*))
  n
)

;; write ((EXPORT . yyc) ...) into the workbook's Library sheet (newest wins)
(defun yyc:lib-save (doc res / book x wb libws lib next k i)
  (if (and res (setq book (yyc:get "LayerBook" nil)) (findfile book) (setq x (yyc:xl-open book nil))
           (setq wb (cadr x) libws (yyc:xl-sheet wb "Library")))
    (progn
      (setq lib (yyc:xl-library wb) next 2)
      (foreach r (yyc:xl-rows libws "A2:A5000") (if (/= (car r) "") (setq next (1+ next))))
      (foreach m res
        (setq k (yyc:strip-num (strcase (car m))))
        (if (setq i (caddr (assoc k lib)))
          (yyc:xl-put libws (strcat "B" (itoa i)) (cdr m))
          (progn
            (yyc:xl-put libws (strcat "A" (itoa next)) k)
            (yyc:xl-put libws (strcat "B" (itoa next)) (cdr m))
            (setq lib (cons (list k (cdr m) next) lib) next (1+ next))))
        (yyc:xl-put libws (strcat "C" (itoa (caddr (assoc k lib)))) (yyc:stamp))
        (yyc:xl-put libws (strcat "D" (itoa (caddr (assoc k lib)))) (yyc:dname doc)))
      (vl-catch-all-apply 'vlax-invoke-method (list wb 'Save))
      T)
  )
)

;; the layers still to walk: ((NAME . objects) ...), most objects first
(defun yyc:walk-todo (doc std cnt / u res)
  (vlax-for L (vla-get-Layers doc)
    (setq u (strcase (vla-get-Name L)))
    (if (and (= (yyc:layer-status u std) "NON-STANDARD") (not (wcmatch u "VIEWPORT#*")))
      (setq res (cons (cons (vla-get-Name L) (cond ((cdr (assoc u cnt))) (0))) res))))
  (vl-sort res '(lambda (a b) (> (cdr a) (cdr b))))
)

(defun yyc:lw-sel () (car (nth (atoi *yyc-wsel*) *yyc-wtodo*)))

(defun yyc:lw-fill-layers ()
  (start_list "lays")
  (foreach x *yyc-wtodo* (add_list (strcat (car x) "   (" (itoa (cdr x)) ")")))
  (end_list)
  (if *yyc-wtodo* (set_tile "lays" *yyc-wsel*))
  (set_tile "count" (strcat (itoa (length *yyc-wtodo*)) " layer(s) left"))
)

;; target list = standard layers matching the filter; preselect pick if it's there
(defun yyc:lw-fill-targets (pick / pat i at)
  (setq pat (strcat "*" (strcase (yyc:trim (get_tile "filt"))) "*") i 0)
  (setq *yyc-wtgts* (vl-remove-if-not '(lambda (n) (wcmatch n pat)) *yyc-wnames*))
  (start_list "tgts")
  (foreach n *yyc-wtgts* (add_list n))
  (end_list)
  (if (and pick (setq at (vl-position (strcase pick) *yyc-wtgts*))) (set_tile "tgts" (itoa at)) (set_tile "tgts" ""))
  (yyc:lw-desc)
)

(defun yyc:lw-desc ( / v r)
  (setq v (get_tile "tgts"))
  (set_tile "desc" (if (and v (/= v "") (setq r (assoc (nth (atoi v) *yyc-wtgts*) *yyc-wstd*)))
                     (strcat "YYC: " (cond ((nth 5 r)) ("")))
                     ""))
)

;; when a layer is picked: filter to its discipline and preselect the best guess
(defun yyc:lw-pick-layer ( / nm sug)
  (if (setq nm (yyc:lw-sel))
    (progn
      (setq sug (car (yyc:suggest (yyc:strip-num (strcase nm)) *yyc-wstd*)))
      (set_tile "filt" (if (wcmatch (strcase nm) "?-*") (substr (strcase nm) 1 2) ""))
      (yyc:lw-fill-targets (if (/= sug "") sug nil))
      (set_tile "msg" (if (/= sug "") (strcat "Best guess: " sug) "No close match - type part of a name in Filter and press Enter.")))
  )
)

(defun yyc:lw-map ( / nm v to n)
  (setq nm (yyc:lw-sel) v (get_tile "tgts"))
  (cond
    ((not nm) (set_tile "msg" "Nothing left to map."))
    ((or (not v) (= v "")) (set_tile "msg" "Pick a YYC layer on the right first."))
    (T
     (setq to (nth (atoi v) *yyc-wtgts*)
           n (yyc:merge-one *yyc-wdoc* nm to))
     (setq *yyc-wdone* (cons (cons nm to) *yyc-wdone*)
           *yyc-wtodo* (vl-remove (assoc nm *yyc-wtodo*) *yyc-wtodo*))
     (if (>= (atoi *yyc-wsel*) (length *yyc-wtodo*)) (setq *yyc-wsel* (itoa (max 0 (1- (length *yyc-wtodo*))))))
     (yyc:lw-fill-layers)
     (yyc:lw-pick-layer)
     (set_tile "msg" (strcat nm " -> " to "  (" (itoa n) " object(s) moved)")))
  )
)

;; see the layer on its own in the current space, then put every layer back
(defun yyc:lw-show (doc nm / ss saved)
  (setq ss (ssget "_X" (list (cons 8 nm) (cons 410 (getvar "CTAB")))))
  (vlax-for L (vla-get-Layers doc)
    (setq saved (cons (cons L (vla-get-LayerOn L)) saved))
    (if (/= (strcase (vla-get-Name L)) (strcase nm)) (vla-put-LayerOn L :vlax-false) (vla-put-LayerOn L :vlax-true)))
  (if ss
    (progn (command "_.ZOOM" "_O" ss "") (sssetfirst nil ss))
    (vla-Regen doc acActiveViewport))
  (yyc:msg (strcat "Showing only " nm (if ss (strcat " - " (itoa (sslength ss)) " object(s) here in " (getvar "CTAB"))
                                         (strcat " - nothing on it in " (getvar "CTAB") " (its objects are in another layout or inside blocks)"))))
  (getstring "\nLook around, then press Enter to go back to the list: ")
  (sssetfirst nil nil)
  (foreach s saved (vla-put-LayerOn (car s) (cdr s)))
  (if ss (command "_.ZOOM" "_P"))
  (vla-Regen doc acActiveViewport)
)

(defun yyc:lw-dcl (path / f)
  (setq f (open path "w"))
  (foreach x (list
    "yycwalk : dialog { label = \"YYC layer walk - map what's left, one layer at a time\";"
    "  : row {"
    "    : column { : text { key = \"count\"; width = 40; }"
    "      : list_box { key = \"lays\"; label = \"Not YYC yet  (objects)\"; width = 42; height = 22; } }"
    "    : column {"
    "      : edit_box { key = \"filt\"; label = \"Filter (Enter):\"; edit_width = 22; }"
    "      : list_box { key = \"tgts\"; label = \"YYC layer it should be\"; width = 34; height = 20; }"
    "    }"
    "  }"
    "  : text { key = \"desc\"; width = 80; }"
    "  : text { key = \"msg\"; width = 80; }"
    "  : row { : button { key = \"show\"; label = \"Show it\"; width = 14; }"
    "          : button { key = \"map\"; label = \"Map to YYC layer\"; is_default = true; }"
    "          : button { key = \"cancel\"; label = \"Done\"; is_cancel = true; width = 14; } }"
    "}") (write-line x f))
  (close f)
)

(defun yyc:t-laywalk (doc / std dcl id res go nm)
  (setq std (yyc:std-layers))
  (if (not std)
    (yyc:msg "No YYC standard in this profile - YYCPROFILES, set Standard layers CSV.")
    (progn
      (setq *yyc-wdoc* doc *yyc-wstd* std *yyc-wdone* nil *yyc-wsel* "0"
            *yyc-wnames* (vl-sort (mapcar 'car std) '<)
            *yyc-wtodo* (yyc:walk-todo doc std (nth 0 (yyc:walk doc))))
      (if (not *yyc-wtodo*)
        (yyc:msg "Every layer is already YYC (or a YYC layer with a modifier). Nothing to walk.")
        (progn
          (setq dcl (vl-filename-mktemp "yycwalk.dcl") go T)
          (yyc:lw-dcl dcl)
          (setq id (load_dialog dcl))
          (while (and go *yyc-wtodo*)
            (if (and (>= id 0) (new_dialog "yycwalk" id))
              (progn
                (yyc:lw-fill-layers)
                (yyc:lw-pick-layer)
                (action_tile "lays" "(setq *yyc-wsel* $value) (yyc:lw-pick-layer) (if (= $reason 4) (done_dialog 3))")
                (action_tile "filt" "(yyc:lw-fill-targets nil)")
                (action_tile "tgts" "(yyc:lw-desc) (if (= $reason 4) (yyc:lw-map))")
                (action_tile "map"  "(yyc:lw-map)")
                (action_tile "show" "(done_dialog 3)")
                (setq res (start_dialog))
                (if (and (= res 3) (setq nm (yyc:lw-sel))) (yyc:lw-show doc nm) (setq go nil)))
              (progn (yyc:msg "Could not open the layer walk window.") (setq go nil))))
          (if (>= id 0) (unload_dialog id))
          (vl-file-delete dcl)
          (foreach m (reverse *yyc-wdone*) (yyc:msg (strcat "  " (car m) " -> " (cdr m))))
          (yyc:msg (strcat "Mapped " (itoa (length *yyc-wdone*)) " layer(s); " (itoa (length *yyc-wtodo*)) " still not YYC."))
          (if *yyc-wdone*
            (if (yyc:lib-save doc *yyc-wdone*)
              (yyc:msg "Remembered in the Layer Mapper workbook's Library - next drawing they come back REMEMBERED.")
              (yyc:msg "Not remembered (no Layer Mapper workbook set in this profile).")))
          (if *yyc-wdone* (yyc:msg "Next: YYCLAYSYNC to give the YYC layers their colour, linetype and lineweight."))
        )
      )
    )
  )
)
(defun c:YYCLAYWALK () (yyc:run 'yyc:t-laywalk))

;; --- layer 0 / DEFPOINTS: report (and select when interactive), never move ---
(defun yyc:zero-objects (doc / res nm)
  (foreach lay (cons (vla-Item (vla-get-Layouts doc) "Model") (yyc:paper-layouts doc))
    (setq nm (vla-get-Name lay))
    (vlax-for o (vla-get-Block lay)
      (if (and (/= (vla-get-ObjectName o) "AcDbViewport")
               (member (strcase (vla-get-Layer o)) '("0" "DEFPOINTS")))
        (setq res (cons (cons nm o) res)))
    )
  )
  res
)

(defun yyc:t-zero (doc / objs by here ss)
  (if (not (setq objs (yyc:zero-objects doc)))
    (yyc:msg "PASS - nothing on 0 or DEFPOINTS outside block definitions.")
    (progn
      (foreach pr objs
        (setq by (yyc:inc (strcat (car pr) " / " (vla-get-Layer (cdr pr)) " / " (vla-get-ObjectName (cdr pr))) by)))
      (yyc:msg (strcat (itoa (length objs)) " object(s) on 0 / DEFPOINTS (layout / layer / type):"))
      (foreach b (vl-sort by '(lambda (x y) (< (car x) (car y)))) (yyc:msg (strcat "  " (car b) ": " (itoa (cdr b)))))
      (if (and (not (yyc:batch-p)) (yyc:active-p doc))
        (progn
          (setq here (vla-get-Name (vla-get-ActiveLayout doc)) ss (ssadd))
          (foreach pr objs (if (= (car pr) here) (ssadd (vlax-vla-object->ename (cdr pr)) ss)))
          (if (> (sslength ss) 0)
            (progn (sssetfirst nil ss)
                   (yyc:msg (strcat (itoa (sslength ss)) " of them selected in " here " - move them to the right layer."))))
        )
      )
    )
  )
)
(defun c:YYCZERO () (yyc:run 'yyc:t-zero))
;; --- YYCPARK / YYCUNPARK: move things out of the way to work on them, then put them back.
;; YYCPARK moves each selected object (e.g. the layer-0 blocks YYCZERO selected) to the left
;; of the drawing, one after another with a gap between them (200000 mm by default). Explode,
;; check, fix layers there. YYCUNPARK moves everything inside each parked spot back by exactly
;; the same distance - exploded pieces included. The spots are saved in the drawing.
(defun yyc:bbox (o / lo hi)
  (if (not (yyc:err-p (vl-catch-all-apply 'vla-GetBoundingBox (list o 'lo 'hi))))
    (list (vlax-safearray->list lo) (vlax-safearray->list hi)))
)

(defun c:YYCPARK ( / doc ss gap own owner minx bb objs cursor dx slots n i o)
  (setq doc (yyc:doc))
  (if (not (setq ss (ssget "_I")))
    (progn (yyc:msg "Select the objects to park (YYCZERO selects what's on layer 0 for you):")
           (setq ss (ssget))))
  (if ss
    (progn
      (setq gap (getdist "\nGap between parked objects and from the drawing <200000>: "))
      (if (not gap) (setq gap 200000.0))
      (setq i 0)
      (repeat (sslength ss) (setq objs (cons (vlax-ename->vla-object (ssname ss i)) objs) i (1+ i)))
      ;; the space (model or a layout) the first object lives in - its block record, by handle
      (setq owner (vlax-ename->vla-object (cdr (assoc 330 (entget (ssname ss 0)))))
            own (vla-get-Handle owner))
      ;; start left of everything else in this space
      (vlax-for o owner
        (if (and (not (member o objs)) (setq bb (yyc:bbox o)))
          (if (or (not minx) (< (car (car bb)) minx)) (setq minx (car (car bb))))))
      (if (not minx) (setq minx 0.0))
      (vla-StartUndoMark doc)
      (setq cursor (- minx gap) n 0
            slots (vl-catch-all-apply 'vlax-ldata-get (list "YYC" "Parked")))
      (if (yyc:err-p slots) (setq slots nil))
      (foreach o (reverse objs)
        (if (and (= (vla-get-Handle (vlax-ename->vla-object (cdr (assoc 330 (entget (vlax-vla-object->ename o)))))) own)
                 (setq bb (yyc:bbox o))
                 (not (yyc:err-p (vl-catch-all-apply 'vla-Move
                        (list o (vlax-3d-point 0 0 0) (vlax-3d-point (setq dx (- cursor (car (cadr bb)))) 0 0))))))
          (progn
            (setq slots (cons (list dx (list (+ (car (car bb)) dx) (cadr (car bb))) (list (+ (car (cadr bb)) dx) (cadr (cadr bb)))
                                    (vla-get-Handle owner)) slots)
                  cursor (- (+ (car (car bb)) dx) gap)
                  n (1+ n)))))
      (vlax-ldata-put "YYC" "Parked" slots)
      (vla-EndUndoMark doc)
      (sssetfirst nil nil)
      (yyc:msg (strcat "Parked " (itoa n) " object(s) to the left of the drawing, " (rtos gap 2 0) " apart."))
      (yyc:msg "Zoom out to find them (ZOOM E). Explode, check, fix their layers there - then run YYCUNPARK to put everything back.")
    )
  )
  (princ)
)

(defun c:YYCUNPARK ( / doc slots owner m lo hi bb c n kill)
  (setq doc (yyc:doc) n 0
        slots (vl-catch-all-apply 'vlax-ldata-get (list "YYC" "Parked")))
  (if (or (yyc:err-p slots) (not slots))
    (yyc:msg "Nothing parked in this drawing.")
    (progn
      (vla-StartUndoMark doc)
      (foreach sl slots
        (setq owner (yyc:try 'vla-HandleToObject (list doc (nth 3 sl))))
        (if owner
          (progn
            (setq lo (nth 1 sl) hi (nth 2 sl)
                  m (max 100.0 (* 0.05 (distance lo hi))))
            (setq kill nil)
            (vlax-for o owner
              (if (setq bb (yyc:bbox o))
                (progn
                  (setq c (list (/ (+ (car (car bb)) (car (cadr bb))) 2.0) (/ (+ (cadr (car bb)) (cadr (cadr bb))) 2.0)))
                  (if (and (<= (- (car lo) m) (car c) (+ (car hi) m)) (<= (- (cadr lo) m) (cadr c) (+ (cadr hi) m)))
                    (setq kill (cons o kill))))))
            (foreach o kill
              (if (not (yyc:err-p (vl-catch-all-apply 'vla-Move (list o (vlax-3d-point (car sl) 0 0) (vlax-3d-point 0 0 0)))))
                (setq n (1+ n)))))))
      (vlax-ldata-delete "YYC" "Parked")
      (vla-EndUndoMark doc)
      (yyc:msg (strcat "Moved " (itoa n) " object(s) back to where they came from (exploded pieces included)."))
    )
  )
  (princ)
)

;;; ---------------------------------------------------------------------------
;;; Drawing walk - one ActiveX pass over every object, shared by QA and LAYEXPORT
;;; Returns (layerCounts blockRefs emptyText zeroTop colTop ltTop lwTop
;;;          colBlk ltBlk lwBlk dims dimsNonAssoc dimsOverride emptyBlocks)
;;; ---------------------------------------------------------------------------

(defun yyc:colmethod (o / tc m)
  (setq tc (vl-catch-all-apply 'vla-get-TrueColor (list o)))
  (if (not (yyc:err-p tc)) (progn (setq m (vla-get-ColorMethod tc)) (vlax-release-object tc)))
  m
)

(defun yyc:dim-assoc-p (o / xd)
  (and (= (vla-get-HasExtensionDictionary o) :vlax-true)
       (setq xd (yyc:try 'vla-GetExtensionDictionary (list o)))
       (yyc:try 'vla-Item (list xd "ACAD_DIMASSOC")))
)

(defun yyc:walk (doc / cnt refs etxt zero colT ltT lwT colB ltB lwB dims dnA dOv eBlk
                       top bname on lay cm lt lw en tov)
  (setq etxt 0 zero 0 colT 0 ltT 0 lwT 0 colB 0 ltB 0 lwB 0 dims 0 dnA 0 dOv 0)
  (vlax-for blk (vla-get-Blocks doc)
    (if (yyc:own-block-p blk)
      (progn
        (setq top (= (vla-get-IsLayout blk) :vlax-true) bname (vla-get-Name blk))
        (if (and (not top) (= (vla-get-Count blk) 0) (/= (substr bname 1 1) "*")) (setq eBlk (cons bname eBlk)))
        (vlax-for o blk
          (setq on (vla-get-ObjectName o) lay (yyc:try 'vla-get-Layer (list o)))
          (if lay (setq lay (strcase lay) cnt (yyc:inc lay cnt)))
          (if (member on '("AcDbBlockReference" "AcDbMInsertBlock"))
            (progn
              (setq refs (yyc:inc (strcase (vla-get-Name o)) refs))
              (if (setq en (yyc:try 'vla-get-EffectiveName (list o))) (setq refs (yyc:inc (strcase en) refs)))
              (if (= (vla-get-HasAttributes o) :vlax-true)
                (foreach a (vlax-invoke o 'GetAttributes) (setq cnt (yyc:inc (strcase (vla-get-Layer a)) cnt))))
            )
          )
          (if (yyc:empty-text-p o) (setq etxt (1+ etxt)))
          (if (/= on "AcDbViewport")
            (progn
              (setq cm (yyc:colmethod o) lt (yyc:try 'vla-get-Linetype (list o)) lw (yyc:try 'vla-get-Lineweight (list o)))
              (if top
                (progn
                  (if (member lay '("0" "DEFPOINTS")) (setq zero (1+ zero)))
                  (if (and cm (/= cm acColorMethodByLayer)) (setq colT (1+ colT)))
                  (if (and lt (/= (strcase lt) "BYLAYER")) (setq ltT (1+ ltT)))
                  (if (and lw (/= lw acLnWtByLayer)) (setq lwT (1+ lwT)))
                )
                (progn
                  (if (and cm (not (member cm (list acColorMethodByLayer acColorMethodByBlock)))) (setq colB (1+ colB)))
                  (if (and lt (not (member (strcase lt) '("BYLAYER" "BYBLOCK")))) (setq ltB (1+ ltB)))
                  (if (and lw (not (member lw (list acLnWtByLayer acLnWtByBlock)))) (setq lwB (1+ lwB)))
                )
              )
            )
          )
          (if (and top (wcmatch on "AcDb*Dimension"))
            (progn
              (setq dims (1+ dims))
              (if (not (yyc:dim-assoc-p o)) (setq dnA (1+ dnA)))
              (setq tov (yyc:try 'vla-get-TextOverride (list o)))
              (if (and tov (/= tov "") (not (vl-string-search "<>" tov))) (setq dOv (1+ dOv)))
            )
          )
        )
      )
    )
  )
  (list cnt refs etxt zero colT ltT lwT colB ltB lwB dims dnA dOv eBlk)
)

;;; ---------------------------------------------------------------------------
;;; Step 9 - QA (CADD Manual 5.15) - report only
;;; ---------------------------------------------------------------------------

(defun yyc:r (s) (setq *yyc-rep* (cons s *yyc-rep*)))

(defun yyc:check (title ok details)
  (yyc:r (strcat (if ok "[PASS] " "[FLAG] ") title))
  (if (not ok) (setq *yyc-flags* (cons title *yyc-flags*)))
  (setq *yyc-res* (cons (list title ok (car details) details) *yyc-res*))
  (foreach x details (yyc:r (strcat "         " x)))
)

;; the summary table: one column per check, in this order, with a short heading
(setq *yyc-qa-cols*
  '(("Purged - no blank layers or unused blocks" . "Purged")
    ("No empty text, no blocks without geometry" . "No empty objects")
    ("Nothing on 0 / DEFPOINTS outside block definitions" . "Layer 0 / Defpoints")
    ("Colour and linetype BY LAYER" . "ByLayer")
    ("Lineweights at Default" . "Lineweights")
    ("Dimensions associative, no text overrides" . "Dimensions")
    ("Linetypes come from CAA.LIN" . "Linetypes")
    ("Text styles use the CAA fonts" . "Fonts")
    ("Layers match the YYC layer standard" . "Layer standard")
    ("Viewports on VIEWPORT# no-plot layers, locked" . "Viewport layers + lock")
    ("Viewports at standard scales" . "Viewport scale")
    ("No empty viewports" . "Empty viewports")
    ("Nothing left behind away from the YYC grid" . "Left at Revit origin")
    ("Main layout active, plots with the YYC page setup" . "Page setup + main layout")
    ("No \"$0$\" names left from binding" . "$0$ names")
    ("Xrefs resolve (eTransmit needs them)" . "Xrefs")))

;; one clean cell: OK, FIX - <short reason>, or blank if the check didn't run
(defun yyc:qa-cell (title / r txt)
  (setq r (assoc title *yyc-res*))
  (cond
    ((not r) "")
    ((cadr r) (if (and (caddr r) (wcmatch (caddr r) "Skipped*")) "skipped" "OK"))
    (T (setq txt (if (caddr r) (caddr r) ""))
       (if (> (strlen txt) 90) (setq txt (strcat (substr txt 1 87) "...")))
       (strcat "FIX - " (vl-string-translate ",\"" ";'" txt))))
)

(defun yyc:first-n (lst n / out)
  (while (and lst (> n 0)) (setq out (cons (car lst) out) lst (cdr lst) n (1- n)))
  (reverse out)
)

(defun yyc:list-line (label lst)
  (if lst
    (list (strcat label " (" (itoa (length lst)) "): " (yyc:join (yyc:first-n lst 25) ", ") (if (> (length lst) 25) ", ..." "")))
  )
)

(defun yyc:lin-names (path / f line pos res)
  (if (and path (setq f (open path "r")))
    (progn
      (while (setq line (read-line f))
        (if (and (= (substr line 1 1) "*") (setq pos (vl-string-search "," line)))
          (setq res (cons (strcase (substr line 2 (1- pos))) res))))
      (close f)
    )
  )
  res
)

(defun yyc:t-qa (doc / w cnt refs layers std blank nonstd diffs lwbad cl curL unused lins ltbad allowed fonts
                   bad main act ctb ctbbad xr h f rpt ff offsc empty e stray lo hi)
  (setq *yyc-rep* '() *yyc-flags* '() *yyc-res* '())
  (yyc:msg "YYC QA - reading the drawing (nothing will be changed)...")
  (setq w (yyc:walk doc) cnt (nth 0 w) refs (nth 1 w) layers (vla-get-Layers doc))
  (setq std (yyc:std-layers) curL (strcase (vla-get-Name (vla-get-ActiveLayer doc))))
  (yyc:r "==================================================================")
  (yyc:r (strcat "YYC QA - CADD Manual 5.15    " (yyc:dname doc) "    " (yyc:stamp)))
  (yyc:r (strcat "Tools v" *yyc-version* ", profile \"" (yyc:profile) "\" - report only, nothing in the drawing was changed"))
  (yyc:r "==================================================================")
  ;; layers
  (vlax-for L layers
    (setq cl (vla-get-Name L))
    (if (not (vl-string-search "|" cl))
      (progn
        (if (and (not (assoc (strcase cl) cnt)) (not (member (strcase cl) (list "0" "DEFPOINTS" curL)))) (setq blank (cons cl blank)))
        (if (/= (vla-get-Lineweight L) acLnWtByLwDefault) (setq lwbad (cons cl lwbad)))
        (if std
          (progn
            (if (and (= (yyc:layer-status cl std) "NON-STANDARD") (not (wcmatch (strcase cl) "VIEWPORT#*"))) (setq nonstd (cons cl nonstd)))
            (if (/= (setq h (yyc:layer-diff L std)) "") (setq diffs (cons (strcat cl ": " h) diffs)))
          )
        )
      )
    )
  )
  ;; blocks with no references
  (vlax-for b (vla-get-Blocks doc)
    (setq cl (vla-get-Name b))
    (if (and (= (vla-get-IsLayout b) :vlax-false) (yyc:own-block-p b)
             (not (member (substr cl 1 1) '("*" "_"))) (not (assoc (strcase cl) refs)))
      (setq unused (cons cl unused)))
  )
  (yyc:check "Purged - no blank layers or unused blocks" (not (or blank unused))
    (append (yyc:list-line "Blank layers" (reverse blank)) (yyc:list-line "Blocks with no references" (reverse unused))
            (if (or blank unused) (list "Fix: YYCCLEAN"))))
  (yyc:check "No empty text, no blocks without geometry" (and (= (nth 2 w) 0) (not (nth 13 w)))
    (append (if (> (nth 2 w) 0) (list (strcat "Empty text objects: " (itoa (nth 2 w)) " - YYCCLEAN deletes them")))
            (yyc:list-line "Empty block definitions" (nth 13 w))))
  (yyc:check "Nothing on 0 / DEFPOINTS outside block definitions" (= (nth 3 w) 0)
    (if (> (nth 3 w) 0) (list (strcat (itoa (nth 3 w)) " object(s) - YYCZERO lists and selects them"))))
  (yyc:check "Colour and linetype BY LAYER" (= 0 (+ (nth 4 w) (nth 5 w) (nth 7 w) (nth 8 w)))
    (list (strcat "Model/paper space: colour overrides " (itoa (nth 4 w)) ", linetype overrides " (itoa (nth 5 w)))
          (strcat "Inside blocks (ByBlock allowed): colour " (itoa (nth 7 w)) ", linetype " (itoa (nth 8 w)))
          "Fix: SETBYLAYER (Ctrl+A first, include blocks)"))
  (yyc:check "Lineweights at Default" (and (not lwbad) (= 0 (+ (nth 6 w) (nth 9 w))))
    (append (yyc:list-line "Layers not at Default" (reverse lwbad))
            (if (> (+ (nth 6 w) (nth 9 w)) 0) (list (strcat "Objects with their own lineweight: " (itoa (+ (nth 6 w) (nth 9 w))))))
            (if lwbad (list "Fix: YYCLWDEFAULT"))))
  (yyc:check "Dimensions associative, no text overrides" (= 0 (+ (nth 11 w) (nth 12 w)))
    (list (strcat "Dimensions: " (itoa (nth 10 w)) ", not associative: " (itoa (nth 11 w)) ", overridden text: " (itoa (nth 12 w)))
          "Fix: DIMREASSOCIATE; clear Text override in Properties"))
  ;; linetypes
  (if (setq lins (yyc:lin-names (yyc:get "LinFile" nil)))
    (progn
      ;; the layer tables also use AutoCAD's own Dashed, Hidden, Phantom... - those are fine
      (foreach r std (if (and (nth 2 r) (not (member (nth 2 r) lins))) (setq lins (cons (nth 2 r) lins))))
      (vlax-for lt (vla-get-Linetypes doc)
        (setq cl (strcase (vla-get-Name lt)))
        (if (and (not (member cl '("BYLAYER" "BYBLOCK" "CONTINUOUS"))) (not (vl-string-search "|" cl)) (not (member cl lins)))
          (setq ltbad (cons (vla-get-Name lt) ltbad))))
      (yyc:check "Linetypes come from CAA.LIN" (not ltbad) (yyc:list-line "Not in CAA.LIN or the YYC layer tables" (reverse ltbad)))
    )
    (yyc:check "Linetypes come from CAA.LIN" T (list "Skipped - set the Linetype file (CAA.LIN) in YYCPROFILES"))
  )
  ;; fonts
  (setq allowed (mapcar '(lambda (x) (strcase (yyc:trim x))) (yyc:split (yyc:get "Fonts" "caa_eng.shx,caa_arch.shx,CAA.SHX") ",")))
  (vlax-for st (vla-get-TextStyles doc)
    (setq cl (vla-get-Name st) ff (vla-get-FontFile st))
    (if (and (/= cl "") (not (vl-string-search "|" cl)))
      (progn
        (setq ff (if (= ff "") "(TrueType)" (strcat (vl-filename-base ff) (cond ((vl-filename-extension ff)) (".shx")))))
        (if (not (member (strcase ff) allowed)) (setq fonts (cons (strcat cl " = " ff) fonts)))
      )
    )
  )
  (yyc:check "Text styles use the CAA fonts" (not fonts) (yyc:list-line "Other fonts" (reverse fonts)))
  ;; layer reference
  (if std
    (yyc:check "Layers match the YYC layer standard" (not (or nonstd diffs))
      (append (yyc:list-line "Not in the standard (modifiers allowed)" (reverse nonstd))
              (yyc:first-n (reverse diffs) 25)
              (if (or nonstd diffs) (list "Fix: YYCLAYXL + YYCLAYMAP, then YYCLAYWALK; properties with YYCLAYSYNC"))))
    (yyc:check "Layers match the YYC layer standard" T (list "Skipped - set the Standard layers CSV in YYCPROFILES"))
  )
  ;; viewports
  (foreach pr (yyc:vp-list doc)
    (setq h (cdr pr) cl (vla-get-Layer h))
    (yyc:r (strcat "         viewport in " (car pr) ": " cl ", scale " (yyc:scale-text h)
                   (if (= (vla-get-DisplayLocked h) :vlax-true) ", locked" ", UNLOCKED")))
    (if (or (not (wcmatch (strcase cl) "VIEWPORT#*"))
            (= (vla-get-Plottable (vla-Item layers cl)) :vlax-true)
            (= (vla-get-DisplayLocked h) :vlax-false))
      (setq bad (cons (car pr) bad)))
    (if (setq ff (yyc:scale-off h))
      (setq offsc (cons (strcat (car pr) " at " (yyc:scale-text h) " (nearest 1:" (rtos ff 2 0) ")") offsc)))
    (if (and (yyc:active-p doc) (setq e (yyc:try 'vlax-vla-object->ename (list h)))
             (not (car (yyc:vp-contents doc e (yyc:vp-geom e) 0.0))))
      (setq empty (cons (car pr) empty)))
  )
  (yyc:check "Viewports at standard scales" (not offsc)
    (append (yyc:list-line "Not a standard scale" (reverse offsc)) (if offsc (list "Fix: YYCVPLOCK offers to snap them"))))
  (yyc:check "No empty viewports" (not empty)
    (append (yyc:list-line "Viewports showing nothing, in layouts" (reverse empty)) (if empty (list "Delete them (left over after YYCSCHEDULES / CHSPACE)"))))
  ;; anything left far from the YYC grid after the move
  (if (and (yyc:aligned-mark doc) (setq h (nth 1 (yyc:align-pts (yyc:profile)))))
    (progn
      (setq stray 0)
      (vlax-for o (vla-get-ModelSpace doc)
        (if (not (yyc:err-p (vl-catch-all-apply 'vla-GetBoundingBox (list o 'lo 'hi))))
          (if (> (distance (list (car h) (cadr h))
                           (list (/ (+ (car (vlax-safearray->list lo)) (car (vlax-safearray->list hi))) 2.0)
                                 (/ (+ (cadr (vlax-safearray->list lo)) (cadr (vlax-safearray->list hi))) 2.0)))
                 3000000.0)
            (setq stray (1+ stray)))))
      (yyc:check "Nothing left behind away from the YYC grid" (= stray 0)
        (if (> stray 0) (list (strcat (itoa stray) " model-space object(s) more than 3 km from the grid - probably left at the Revit origin. ZOOM E to find them."))))
    )
  )
  (yyc:check "Viewports on VIEWPORT# no-plot layers, locked" (not bad)
    (append (yyc:list-line "Problem viewports in layouts" (reverse bad))
            (if bad (list "Fix: YYCVPLAYERS, check scale, then YYCVPLOCK") (list "Confirm the scales above are the intended ones"))))
  ;; layouts / page setup
  (setq ctb (strcase (yyc:get "CTB" "YYC_BW_HPv1.ctb")) main (car (yyc:paper-layouts doc))
        act (vla-get-Name (vla-get-ActiveLayout doc)))
  (foreach lay (yyc:paper-layouts doc)
    (yyc:r (strcat "         layout " (vla-get-Name lay) ": " (vla-get-CanonicalMediaName lay) ", " (vla-get-StyleSheet lay)))
    (if (/= (strcase (vla-get-StyleSheet lay)) ctb) (setq ctbbad (cons (vla-get-Name lay) ctbbad)))
  )
  (yyc:check "Main layout active, plots with the YYC page setup"
    (and main (= (strcase act) (strcase (vla-get-Name main))) (not ctbbad))
    (append (if (and main (/= (strcase act) (strcase (vla-get-Name main)))) (list (strcat "Active tab is " act " - main layout is " (vla-get-Name main) " (YYCFINAL)")))
            (yyc:list-line (strcat "Layouts not using " ctb) (reverse ctbbad))
            (if ctbbad (list "Fix: YYCPAGESETUP"))))
  ;; $0$
  (setq h (yyc:find-$0$ doc))
  (yyc:check "No \"$0$\" names left from binding" (not h) (mapcar '(lambda (x) (strcat (car x) ": " (yyc:join (cdr x) ", "))) h))
  ;; xrefs
  (vlax-for b (vla-get-Blocks doc)
    (if (= (vla-get-IsXRef b) :vlax-true)
      (setq xr (cons (strcat (vla-get-Name b) (if (findfile (vla-get-Path b)) "" "  (FILE NOT FOUND)")) xr))))
  (yyc:check "Xrefs resolve (eTransmit needs them)" (not (vl-some '(lambda (x) (vl-string-search "NOT FOUND" x)) xr))
    (if xr (yyc:list-line "Xrefs" (reverse xr)) (list "No xrefs")))
  (yyc:r "[LOOK] By eye: AUDIT clean, ZOOM Extents shows nothing off-sheet,")
  (yyc:r "       linetypes are not drawn dashes, circles are not polyline approximations.")
  (yyc:r "------------------------------------------------------------------")
  (yyc:r (strcat "RESULT: " (if *yyc-flags* (strcat (itoa (length *yyc-flags*)) " check(s) flagged") "all automatic checks passed")))
  (setq *yyc-rep* (reverse *yyc-rep*))
  (if (not (yyc:batch-p)) (foreach x *yyc-rep* (yyc:msg x)))
  (setq rpt (strcat (yyc:dfolder doc) (yyc:dbase doc) "_YYC-QA.txt"))
  (if (setq f (open rpt "w"))
    (progn (foreach x *yyc-rep* (write-line x f)) (close f) (yyc:msg (strcat "Report: " rpt)))
    (yyc:msg "Could not write the report file (folder read-only, or drawing never saved?).")
  )
  (yyc:msg (strcat "QA: " (if *yyc-flags* (strcat (itoa (length *yyc-flags*)) " check(s) to fix") "all automatic checks passed")))
)
;; YYC_QA_Summary.csv - one row per drawing, one column per check. Rows for other
;; drawings are kept; an old-format file is replaced.
(defun yyc:qa-summary (doc / csv header f first me row)
  (setq csv (strcat (yyc:dfolder doc) "YYC_QA_Summary.csv") me (yyc:dname doc))
  (setq header (yyc:join (append (list "Drawing" "Checked" "Result") (mapcar 'cdr *yyc-qa-cols*)) ","))
  (if (and (setq f (open csv "r")) (setq first (read-line f))) (close f) (if f (close f)))
  (if (and first (/= first header)) (vl-file-delete csv))
  (setq row (yyc:join (append (list (yyc:q me) (yyc:stamp)
                                    (if *yyc-flags* (strcat (itoa (length *yyc-flags*)) " to fix") "PASS"))
                              (mapcar '(lambda (c) (yyc:q (yyc:qa-cell (car c)))) *yyc-qa-cols*)) ","))
  (yyc:csv-replace csv me header (list row))
)

;; --- the QA report: YYC QA Report.xlsx (kit template) -> a dated copy next to the drawings
(defun yyc:qa-book ( / b)
  (cond ((and (setq b (yyc:get "QABook" nil)) (findfile b)) b)
        ((setq b (car (yyc:find-all (yyc:get "Kit" nil) '("*QA*Report*.xlsx") 2))) (yyc:set "QABook" b) b)))

(defun yyc:qa-row (doc) (list (yyc:dname doc) (yyc:stamp) (length *yyc-flags*) *yyc-res*))

(defun yyc:qa-status (r)
  (cond ((not r) "")
        ((and (cadr r) (caddr r) (wcmatch (caddr r) "Skipped*")) "skipped")
        ((cadr r) "OK")
        (T "FIX")))

(defun yyc:xl-col (n / s) (setq s "") (while (> n 0) (setq s (strcat (chr (+ 65 (rem (1- n) 26))) s) n (/ (1- n) 26))) s)

;; one row of values into a range in a single call
(defun yyc:xl-row (ws addr vals / arr i)
  (setq arr (vlax-make-safearray vlax-vbVariant (cons 1 1) (cons 1 (length vals))) i 0)
  (foreach v vals (setq i (1+ i)) (vlax-safearray-put-element arr 1 i (vlax-make-variant v)))
  (vlax-put-property (vlax-get-property ws 'Range addr) 'Value2 (vlax-make-variant arr))
)

(defun yyc:qa-report (rows folder / book x wb sm dt out n m r st found lastc)
  (cond
    ((not (setq book (yyc:qa-book)))
     (yyc:msg "No QA report template - put \"YYC QA Report.xlsx\" in the kit folder (or set it in YYCPROFILES). Text reports are next to the drawings.") nil)
    ((not (setq x (yyc:xl-open book nil))) nil)
    (T
     (setq wb (cadr x) out (strcat folder "\\YYC QA Report " (yyc:stamp-file) ".xlsx"))
     (if (yyc:err-p (vl-catch-all-apply 'vlax-invoke-method (list wb 'SaveAs out 51)))
       (progn (yyc:msg (strcat "Could not save " out)) nil)
       (progn
         (setq sm (yyc:xl-sheet wb "Summary") dt (yyc:xl-sheet wb "Details")
               lastc (yyc:xl-col (+ 4 (length *yyc-qa-cols*))))
         (vlax-invoke-method (vlax-get-property sm 'Range "A2:C2000") 'ClearContents)
         (vlax-invoke-method (vlax-get-property sm 'Range (strcat "E2:" lastc "2000")) 'ClearContents)
         (vlax-invoke-method (vlax-get-property dt 'Range "A2:D20000") 'ClearContents)
         (setq n 1 m 1)
         (foreach row rows
           (setq n (1+ n))
           (yyc:xl-row sm (strcat "A" (itoa n) ":C" (itoa n))
                       (list (car row) (cadr row) (if (= (caddr row) 0) "PASS" (strcat "FIX - " (itoa (caddr row))))))
           (yyc:xl-row sm (strcat "E" (itoa n) ":" lastc (itoa n))
                       (mapcar '(lambda (c) (yyc:qa-status (assoc (car c) (cadddr row)))) *yyc-qa-cols*))
           (foreach c *yyc-qa-cols*
             (setq r (assoc (car c) (cadddr row)) st (yyc:qa-status r) m (1+ m))
             (setq found (if r (yyc:join (vl-remove-if '(lambda (s) (wcmatch s "Fix:*")) (nth 3 r)) "; ") "did not run"))
             (if (= found "") (setq found (if (= st "OK") "Nothing found" "")))
             (if (> (strlen found) 900) (setq found (strcat (substr found 1 897) "...")))
             (yyc:xl-row dt (strcat "A" (itoa m) ":D" (itoa m)) (list (car row) (cdr c) st found))))
         (vl-catch-all-apply 'vlax-invoke-method (list wb 'Save))
         (vl-catch-all-apply 'vlax-invoke-method (list sm 'Activate))
         (vlax-put-property (car x) 'Visible :vlax-true)
         (yyc:msg (strcat "QA report: " out))
         (yyc:msg "  Summary = one row per drawing. Details = what was found + how to fix. Filter Status = FIX for your to-do list.")
         out)))
  )
)

;; YYCQA: this drawing, or every drawing in a folder (opened read-only, nothing saved)
(defun c:YYCQA ( / k doc folder files docs d path rows opened od)
  (initget "This Folder")
  (setq k (getkword "\nQA [This drawing/Folder of drawings] <This>: "))
  (if (= k "Folder")
    (if (setq folder (yyc:browse-folder "Folder of drawings to QA"))
      (progn
        (setq files (yyc:dwgs-in folder) docs (vla-get-Documents (yyc:acad)))
        (yyc:msg (strcat "QA on " (itoa (length files)) " drawing(s) in " folder " - read-only, nothing is saved."))
        (setq *yyc-batch* T)
        (foreach x files
          (setq path (strcat folder "\\" x) *yyc-cur* x opened nil)
          (cond
            ((setq od (yyc:open-doc path)) (setq d od))
            ((yyc:err-p (setq d (vl-catch-all-apply 'vla-Open (list docs path :vlax-true))))
             (setq d nil) (yyc:msg "Could not open - skipped."))
            (T (setq opened T)))
          (if d
            (progn
              (if (yyc:err-p (setq k (vl-catch-all-apply 'yyc:t-qa (list d))))
                (yyc:msg (strcat "ERROR: " (vl-catch-all-error-message k)))
                (setq rows (cons (yyc:qa-row d) rows)))
              (if opened (vl-catch-all-apply 'vla-Close (list d :vlax-false)))
              (yyc:msg (strcat (itoa (length *yyc-flags*)) " to fix")))))
        (setq *yyc-batch* nil *yyc-cur* nil)
        (if rows (yyc:qa-report (reverse rows) folder))))
    (progn
      (setq doc (yyc:doc))
      (yyc:run 'yyc:t-qa)
      (if *yyc-res* (yyc:qa-report (list (yyc:qa-row doc)) (vl-string-right-trim "\\" (yyc:dfolder doc))))))
  (princ)
)

;;; ---------------------------------------------------------------------------
;;; Step 9 - Final: lock, main layout active, zoom extents
;;; ---------------------------------------------------------------------------

(defun yyc:t-final (doc / main n)
  (setq n (yyc:t-vplock doc) main (car (yyc:paper-layouts doc)))
  (if (yyc:active-p doc)
    (progn
      (vla-put-ActiveLayout doc (vla-Item (vla-get-Layouts doc) "Model"))
      (vla-ZoomExtents (yyc:acad))
    )
  )
  (if main
    (progn
      (vla-put-ActiveLayout doc main)
      (if (yyc:active-p doc)
        (progn (vl-catch-all-apply 'vla-put-MSpace (list doc :vlax-false)) (vla-ZoomExtents (yyc:acad))))
      (yyc:msg (strcat "Main layout \"" (vla-get-Name main) "\" is active"
                       (if (yyc:active-p doc) ", zoomed to extents." " (zoom skipped - drawing not on screen).")))
    )
    (yyc:msg "No paper-space layout found.")
  )
  (if (not (yyc:batch-p)) (yyc:msg "Look for anything stranded off the sheet, then save."))
)
(defun c:YYCFINAL () (yyc:run 'yyc:t-final))
;;; ---------------------------------------------------------------------------
;;; Step 10 - file list for the File Description
;;; ---------------------------------------------------------------------------

(defun yyc:name-ok (base dwgno)
  (and (= (substr (strcase base) 1 (strlen dwgno)) (strcase dwgno))
       (vl-every '(lambda (c) (or (<= 48 c 57) (<= 65 c 90))) (vl-string->list (strcase base))))
)

;; Sheet title from a Revit titleblock block name, e.g.
;; "YYC_A0_WD - A0 - Titleblock-29504745-DEPARTURES OVERALL FLOOR PLAN"
;; -> everything after the last all-digits piece (Revit's element id).
(defun yyc:title-from-name (nm / parts i k)
  (setq parts (yyc:split nm "-") i 0 k nil)
  (foreach x parts
    (if (and (> (strlen (yyc:trim x)) 3) (vl-every '(lambda (c) (<= 48 c 57)) (vl-string->list (yyc:trim x)))) (setq k i))
    (setq i (1+ i)))
  (if (and k (< (1+ k) (length parts)))
    (yyc:trim (yyc:join (cdr (member (nth k parts) (vl-remove-if '(lambda (x) nil) parts))) "-")))
)

;; sheet title of a closed drawing (ObjectDBX): the YYC titleblock's SHEET-TITLE
;; fields if it has them, otherwise the title in the Revit titleblock's block name
;; the drawing if it's open in AutoCAD (ObjectDBX can't read an open file)
(defun yyc:open-doc (path / hit)
  (vlax-for d (vla-get-Documents (yyc:acad))
    (if (= (strcase (vla-get-FullName d)) (strcase path)) (setq hit d)))
  hit
)

(defun yyc:read-title (path / dbx title parts nm odoc)
  (setq odoc (yyc:open-doc path))
  (if (setq dbx (if odoc odoc (yyc:dbx-open path)))
    (progn
      (foreach lay (yyc:paper-layouts dbx)
        (if (not title)
          (vlax-for o (vla-get-Block lay)
            (if (and (not title) (= (vla-get-ObjectName o) "AcDbBlockReference")
                     (wcmatch (strcase (setq nm (vla-get-Name o))) "*TITLEBLOCK*"))
              (progn
                (setq parts nil)
                (if (= (vla-get-HasAttributes o) :vlax-true)
                  (foreach a (vlax-invoke o 'GetAttributes)
                    (if (wcmatch (strcase (vla-get-TagString a)) "SHEET-TITLE-#")
                      (setq parts (cons (cons (vla-get-TagString a) (yyc:trim (vla-get-TextString a))) parts)))))
                (setq parts (vl-remove-if '(lambda (x) (member (strcase (cdr x)) '("" "DRAWING NAME 1" "DRAWING NAME 2" "DRAWING NAME 3")))
                              (vl-sort parts '(lambda (x y) (< (car x) (car y))))))
                (setq title (cond (parts (strcase (yyc:join (mapcar 'cdr parts) " ")))
                                  ((yyc:title-from-name nm) (strcase (yyc:title-from-name nm))))))))))
      ;; no titleblock title: any Revit block on the sheet named "...-<id>-<title>" (e.g. the view title)
      (if (not title)
        (foreach lay (yyc:paper-layouts dbx)
          (vlax-for o (vla-get-Block lay)
            (if (and (not title) (= (vla-get-ObjectName o) "AcDbBlockReference"))
              (setq title (yyc:title-from-name (vla-get-Name o)))))))
      (if title (setq title (strcase title)))
      (if (not odoc) (yyc:dbx-close dbx))
    )
  )
  title
)

(defun c:YYCFILELIST ( / folder files dwgno f bad lines disc title cur notitle)
  (setq folder (yyc:browse-folder "Delivery folder to list"))
  (if folder
    (progn
      (setq dwgno (strcase (yyc:ask "YYC drawing number" (yyc:get "DwgNo" "24C024"))))
      (setq files (yyc:dwgs-in folder))
      (if (setq f (open (strcat folder "\\filelist.txt") "w"))
        (progn (foreach x files (write-line x f)) (close f)))
      (foreach x files (if (not (yyc:name-ok (vl-filename-base x) dwgno)) (setq bad (cons x bad))))
      (yyc:msg (strcat (itoa (length files)) " drawing(s) written to " folder "\\filelist.txt"))
      ;; File Description lines from each drawing's YYC titleblock, grouped by discipline letter
      (yyc:msg "Reading sheet titles from the titleblocks...")
      (foreach x files
        (setq title (yyc:read-title (strcat folder "\\" x))
              disc (substr (strcase (vl-filename-base x)) (1+ (strlen dwgno)) 1))
        (if (not title) (setq notitle (cons x notitle)))
        (setq lines (cons (list disc x (if title title "(TITLE NOT FOUND)")) lines)))
      (setq lines (vl-sort lines '(lambda (a b) (if (= (car a) (car b)) (< (cadr a) (cadr b)) (< (car a) (car b))))))
      (if (setq f (open (strcat folder "\\FileDescription_list.txt") "w"))
        (progn
          (foreach l lines
            (if (/= (car l) cur)
              (progn (setq cur (car l)) (write-line "" f)
                     (write-line (cdr (cond ((assoc cur '(("A" . "ARCHITECTURAL") ("E" . "ELECTRICAL") ("M" . "MECHANICAL") ("S" . "STRUCTURAL")
                                                         ("C" . "CIVIL") ("F" . "FIRE PROTECTION") ("G" . "GENERAL") ("I" . "INTERIORS")
                                                         ("L" . "LANDSCAPE") ("P" . "PLUMBING") ("T" . "TELECOM"))))
                                            ((cons cur cur)))) f)))
            (write-line (strcat (cadr l) " - " (caddr l)) f))
          (close f)
          (yyc:msg (strcat "File Description lines written to " folder "\\FileDescription_list.txt - paste them into the File Description."))
          (if notitle (yyc:msg (strcat "No sheet title found in: " (yyc:join (reverse notitle) ", ") " - type those by hand.")))))
      (if bad
        (progn
          (yyc:msg (strcat "Names that don't follow " dwgno " + sheet number (letters and digits only):"))
          (foreach x (reverse bad) (yyc:msg (strcat "  " x)))
          (yyc:msg "Fix: YYCRENAME, or rename by hand.")
        )
        (yyc:msg "All file names follow the pattern.")
      )
    )
  )
  (princ)
)


;; --- YYCFILEDESC: one File Description for the whole delivery, Word + PDF -------------
;; Sources (pick as many as you like): each consultant's list - a FileDescription_list.txt
;; from YYCFILELIST, their own .docx or .txt/.csv list - or a folder of their DWGs (titles
;; read from the titleblocks). Any line "name.dwg - TITLE" counts. The template is your
;; last File Description .docx: its header is kept, today's date goes in, and the file list
;; is replaced with the combined one, grouped by discipline (A, S, M, E, then the rest).
(defun yyc:fd-parse (line / u pos name rest)
  (setq u (strcase line))
  (if (setq pos (vl-string-search ".DWG" u))
    (progn
      (setq name (yyc:trim (substr line 1 (+ pos 4)))
            rest (yyc:trim (substr line (+ pos 5))))
      (setq rest (vl-string-left-trim " ,;\t-\226\227" rest))  ; - , ; tab, en/em dash
      (if (and (> (strlen name) 4) (not (vl-string-search " " name)))
        (cons name (yyc:trim rest))))))

(defun yyc:fd-word ( / w) (setq w (vl-catch-all-apply 'vlax-get-or-create-object (list "Word.Application"))) (if (yyc:err-p w) nil w))

(defun yyc:fd-read (path / ext f line res wd d txt)
  (setq ext (strcase (cond ((vl-filename-extension path)) (""))))
  (cond
    ((member ext '(".DOCX" ".DOC"))
     (if (and (setq wd (yyc:fd-word))
              (not (yyc:err-p (setq d (vl-catch-all-apply 'vlax-invoke-method (list (vlax-get-property wd 'Documents) 'Open path :vlax-false :vlax-true))))))
       (progn
         (setq txt (vlax-get-property (vlax-get-property d 'Content) 'Text))
         (vl-catch-all-apply 'vlax-invoke-method (list d 'Close :vlax-false))
         (foreach l (yyc:split txt "\r") (if (setq l (yyc:fd-parse l)) (setq res (cons l res)))))
       (yyc:msg (strcat "Could not read " path " with Word."))))
    (T
     (if (setq f (open path "r"))
       (progn (while (setq line (read-line f)) (if (setq line (yyc:fd-parse line)) (setq res (cons line res)))) (close f)))))
  (reverse res)
)

(defun yyc:fd-disc (name dwgno / b)
  (setq b (strcase (vl-filename-base name)))
  (if (= (substr b 1 (strlen dwgno)) (strcase dwgno)) (substr b (1+ (strlen dwgno)) 1) (substr b 1 1))
)

;; header values from a delivered drawing's titleblock: the value is the bigger text just
;; under each label (PROJECT 4391, CONTRACT NO. 551, CAA DRAWING NUMBER 24Z017)
(defun yyc:tb-texts (dbx / res nm s ip)
  (foreach lay (yyc:paper-layouts dbx)
    (vlax-for o (vla-get-Block lay)
      (if (and (not res) (= (vla-get-ObjectName o) "AcDbBlockReference")
               (wcmatch (strcase (setq nm (vla-get-Name o))) "*TITLEBLOCK*"))
        (progn
          (vlax-for e (vla-Item (vla-get-Blocks dbx) nm)
            (if (member (vla-get-ObjectName e) '("AcDbText" "AcDbMText"))
              (progn
                (setq s (yyc:trim (vl-string-trim "{}" (vla-get-TextString e))) ip (vlax-get e 'InsertionPoint))
                (setq res (cons (list (strcase s) (car ip) (cadr ip) (vla-get-Height e)) res)))))
          (if (= (vla-get-HasAttributes o) :vlax-true)
            (foreach a (vlax-invoke o 'GetAttributes)
              (setq ip (vlax-get a 'InsertionPoint))
              (setq res (cons (list (strcase (yyc:trim (vla-get-TextString a))) (car ip) (cadr ip) (vla-get-Height a)) res))))))))
  res
)
(defun yyc:tb-value (texts label / best d)
  (foreach lb texts
    (if (= (car lb) label)
      (foreach v texts
        (if (and (/= (car v) "") (> (nth 3 v) (nth 3 lb))
                 (< (nth 2 v) (nth 2 lb)) (< (- (nth 2 lb) (nth 2 v)) 12.0) (< (abs (- (nth 1 v) (nth 1 lb))) 25.0))
          (progn (setq d (distance (cdr lb) (cdr v)))
                 (if (or (not best) (< d (car best))) (setq best (list d (car v)))))))))
  (cadr best)
)
(defun yyc:tb-fields (path / dbx od tx res v)
  (setq od (yyc:open-doc path))
  (if (setq dbx (if od od (yyc:dbx-open path)))
    (progn
      (setq tx (yyc:tb-texts dbx))
      (if (not od) (yyc:dbx-close dbx))
      (foreach k '(("PROJECT" . "[Firm project number]") ("CONTRACT NO." . "[Contract number]") ("CAA DRAWING NUMBER" . "[YYC drawing number]"))
        (if (setq v (yyc:tb-value tx (car k))) (setq res (cons (cons (cdr k) v) res))))))
  res
)

(defun c:YYCFILEDESC ( / dwgno tpl all src k items d n lst order out wd docs doc paras i p start rng txt grp cur stamp pdf miss dwgs nodesc nodwg tbf)
  (setq dwgno (yyc:get "DwgNo" "24C024"))
  (yyc:msg "Combine every consultant's file list into one File Description (Word + PDF).")
  ;; 1. sources
  (while (progn (initget "File Lists Folder Done")
                (setq k (getkword (strcat "\nAdd a consultant list [File/Lists folder (every .txt .docx .csv in it)/Folder of DWGs/Done] <"
                                          (if all "Done" "File") ">: ")))
                (if (not k) (setq k (if all "Done" "File")))
                (/= k "Done"))
    (cond
      ((= k "File")
       (if (setq src (getfiled "Consultant file list (.txt, .docx, .csv)" (yyc:get "FDLast" "") "*" 0))
         (progn (yyc:set "FDLast" (strcat (vl-filename-directory src) "\\"))
                (setq items (yyc:fd-read src))
                (yyc:msg (strcat "  " (itoa (length items)) " drawing(s) from " (vl-filename-base src) (cond ((vl-filename-extension src)) (""))))
                (setq all (append all items)))))
      ((= k "Lists")
       (if (setq src (yyc:browse-folder "Folder with the consultants' lists (.txt .docx .csv)"))
         (foreach x (append (vl-directory-files src "*.txt" 1) (vl-directory-files src "*.docx" 1) (vl-directory-files src "*.csv" 1))
           (if (not (wcmatch x "~$*,filelist.txt"))
             (progn
               (setq items (yyc:fd-read (strcat src "\\" x)))
               (yyc:msg (strcat "  " (itoa (length items)) " drawing(s) from " x))
               (setq all (append all items)))))))
      ((= k "Folder")
       (if (setq src (yyc:browse-folder "Folder of a consultant's DWGs"))
         (progn
           (setq items nil)
           (yyc:msg "  Reading sheet titles from the titleblocks...")
           (foreach x (yyc:dwgs-in src)
             (setq items (cons (cons x (cond ((yyc:read-title (strcat src "\\" x))) ("(TITLE NOT FOUND)"))) items)))
           (yyc:msg (strcat "  " (itoa (length items)) " drawing(s) from " src))
           (setq all (append all (reverse items))))))))
  ;; 2. combine: one line per file name (a later list wins), grouped by discipline
  (foreach it all
    (if (setq d (assoc (strcase (car it)) lst)) (setq lst (subst (list (strcase (car it)) (car it) (cdr it)) d lst))
      (setq lst (cons (list (strcase (car it)) (car it) (cdr it)) lst))))
  (setq order '("A" "S" "M" "E"))
  (setq lst (vl-sort lst '(lambda (a b / da db ia ib)
              (setq da (yyc:fd-disc (cadr a) dwgno) db (yyc:fd-disc (cadr b) dwgno)
                    ia (cond ((vl-position da order)) (99)) ib (cond ((vl-position db order)) (99)))
              (cond ((/= ia ib) (< ia ib)) ((/= da db) (< da db)) (T (< (car a) (car b)))))))
  (cond
    ((not lst) (yyc:msg "No drawings found in those lists."))
    ((not (setq tpl (yyc:need-file "FDTemplate" "File Description template (.docx): the kit's template, or your last File Description" "docx"))) nil)
    ((not (setq wd (yyc:fd-word))) (yyc:msg "Word could not be started on this machine."))
    ((not (setq out (yyc:browse-folder "Where to save the File Description (the delivery folder)"))) nil)
    (T
     (setq stamp (menucmd "M=$(edtime,$(getvar,DATE),YYYY/MO/DD)"))
     (setq docs (vlax-get-property wd 'Documents)
           doc (vlax-invoke-method docs 'Add tpl))
     ;; today's date in the header (first yyyy/mm/dd in the document)
     (setq rng (vlax-get-property doc 'Content))
     (vl-catch-all-apply 'vlax-invoke-method
       (list (vlax-get-property rng 'Find) 'Execute "[0-9]{4}/[0-9]{2}/[0-9]{2}" :vlax-false :vlax-false :vlax-true
             :vlax-false :vlax-false :vlax-true 0 :vlax-false stamp 1))
     ;; header numbers from a delivered drawing's titleblock (project, contract, CAA drawing no.)
     (setq dwgs (vl-remove-if-not '(lambda (x) (wcmatch (strcase x) (strcat (strcase dwgno) "*"))) (yyc:dwgs-in out)))
     (setq tbf (if dwgs (yyc:tb-fields (strcat out "\\" (car dwgs)))))
     (if (not (assoc "[YYC drawing number]" tbf)) (setq tbf (cons (cons "[YYC drawing number]" dwgno) tbf)))
     (if tbf (yyc:msg (strcat "  From the titleblock of " (if dwgs (car dwgs) "the profile") ": "
                              (yyc:join (mapcar '(lambda (f) (strcat (vl-string-trim "[]" (car f)) " " (cdr f))) tbf) ", "))))
     (foreach f tbf
       (vl-catch-all-apply 'vlax-invoke-method
         (list (vlax-get-property (vlax-get-property doc 'Content) 'Find) 'Execute (car f) :vlax-false :vlax-false :vlax-false
               :vlax-false :vlax-false :vlax-true 0 :vlax-false (cdr f) 2)))
     ;; find "File Description:" and replace everything after it
     (setq paras (vlax-get-property doc 'Paragraphs) n (vlax-get-property paras 'Count) i 1 start nil)
     (while (and (<= i n) (not start))
       (setq p (vlax-invoke-method paras 'Item i)
             txt (strcase (yyc:trim (vlax-get-property (vlax-get-property p 'Range) 'Text))))
       (if (wcmatch txt "FILE DESCRIPTION:*,FILE DESCRIPTION") (setq start i))
       (setq i (1+ i)))
     (if (not start)
       (progn (yyc:msg "The template has no \"File Description:\" line - the list is added at the end.") (setq start n)))
     (setq rng (vlax-get-property (vlax-invoke-method paras 'Item (min (1+ start) n)) 'Range))
     (if (< start n)
       (vlax-put-property rng 'End (vlax-get-property (vlax-get-property doc 'Content) 'End))
       (progn (vlax-invoke-method rng 'InsertParagraphAfter) (setq rng (vlax-get-property (vlax-invoke-method paras 'Item (1+ n)) 'Range))))
     (setq txt "" cur nil)
     (foreach l lst
       (setq grp (yyc:fd-disc (cadr l) dwgno))
       (if (and cur (/= grp cur)) (setq txt (strcat txt "\r")))
       (setq cur grp txt (strcat txt (cadr l) " - " (strcase (caddr l)) "\r")))
     (vlax-put-property rng 'Text txt)
     (setq out (strcat out "\\" dwgno "_File Description " (menucmd "M=$(edtime,$(getvar,DATE),YYYY-MO-DD)"))
           pdf (strcat out ".pdf"))
     (if (yyc:err-p (vl-catch-all-apply 'vlax-invoke-method (list doc 'SaveAs2 (strcat out ".docx") 16)))
       (yyc:msg "Could not save the Word file (is one with that name open?).")
       (progn
         (if (yyc:err-p (vl-catch-all-apply 'vlax-invoke-method (list doc 'ExportAsFixedFormat pdf 17)))
           (yyc:msg "Saved the Word file, but the PDF export failed - File > Save As PDF in Word.")
           (yyc:msg (strcat "PDF: " pdf)))
         (yyc:msg (strcat "Word: " out ".docx"))))
     (vlax-put-property wd 'Visible :vlax-true)
     (setq miss (vl-remove-if-not '(lambda (l) (wcmatch (strcase (caddr l)) "*NOT FOUND*,")) lst))
     (yyc:msg (strcat (itoa (length lst)) " drawing(s) listed, grouped A / S / M / E / others."))
     ;; against the delivery folder: every DWG there should have a line, every line a DWG
     (setq dwgs (vl-remove-if-not '(lambda (x) (wcmatch (strcase x) (strcat (strcase dwgno) "*"))) (yyc:dwgs-in (vl-filename-directory pdf))))
     (setq nodesc (vl-remove-if '(lambda (x) (assoc (strcase x) lst)) dwgs)
           nodwg (vl-remove-if '(lambda (l) (member (car l) (mapcar 'strcase dwgs))) lst))
     (if (and dwgs nodesc) (yyc:msg (strcat "  STILL MISSING a description (no consultant list has them yet): " (yyc:join nodesc ", "))))
     (if (and dwgs nodwg) (yyc:msg (strcat "  Listed, but no DWG with that name in the delivery folder: " (yyc:join (mapcar 'cadr nodwg) ", "))))
     (if (and dwgs (not nodesc) (not nodwg)) (yyc:msg "  Every DWG in the delivery folder has a line, and every line has a DWG."))
     (if miss (yyc:msg (strcat "  No title for: " (yyc:join (mapcar 'cadr miss) ", ") " - type them in Word, then save the PDF again.")))
     (if (vl-string-search "[" (vlax-get-property (vlax-get-property doc 'Content) 'Text))
       (progn
         (yyc:msg "The header still has [bracketed] fields - fill them in Word, save, and Save As PDF again.")
         (yyc:msg "Tip: fill them once in the kit's \"YYC File Description - Template.docx\" and they're done for every delivery."))
       (yyc:msg "Check the header in Word (project, phase, stage) - it comes from your template.")))
  )
  (princ)
)

;; --- YYCTRANSMIT: one zip per discipline, named <drawing no>_<stage>_<phase>_<discipline>.zip
;; The discipline comes from the letter after the drawing number (24C024A201 -> A -> AR).
;; Each zip holds the drawings plus what eTransmit would add: xrefs (re-pathed to sit next
;; to the drawing, NOT bound), the fonts the text styles use, CAA.SHX, and the CTB the
;; layouts plot with, plus a Transmittal.txt listing everything. Originals are not touched.
(setq *yyc-zip-codes* '(("A" . "AR") ("E" . "EL") ("M" . "ME") ("S" . "ST")
                        ("C" . "CV") ("F" . "FP") ("G" . "GN") ("I" . "ID") ("L" . "LA") ("P" . "PL") ("T" . "TC") ("V" . "SV")))

(defun yyc:ps-q (p) (strcat "'" (vl-string-subst "''" "'" p) "'"))
(defun yyc:add-uniq (x lst) (if (and x (not (member (strcase x) (mapcar 'strcase lst)))) (append lst (list x)) lst))

;; what one drawing needs: (xref-paths font-paths ctb-paths)
(defun yyc:tx-deps (path / dbx dir xr fo ct f prefs ctbdir)
  (setq dir (vl-filename-directory path)
        prefs (vla-get-Files (vla-get-Preferences (yyc:acad)))
        ctbdir (yyc:try 'vla-get-PrinterStyleSheetPath (list prefs)))
  (if (setq dbx (yyc:dbx-open path))
    (progn
      (vlax-for b (vla-get-Blocks dbx)
        (if (= (vla-get-IsXRef b) :vlax-true)
          (if (setq f (cond ((findfile (vla-get-Path b)))
                            ((findfile (strcat dir "\\" (vl-filename-base (vla-get-Path b)) ".dwg")))))
            (setq xr (yyc:add-uniq f xr)))))
      (vlax-for st (vla-get-TextStyles dbx)
        (foreach ff (list (vla-get-FontFile st) (yyc:try 'vla-get-BigFontFile (list st)))
          (if (and ff (/= ff "") (not (wcmatch (strcase ff) "*.TTF,*.OTF,*.TTC")))
            (if (setq f (cond ((findfile ff)) ((findfile (strcat ff ".shx"))))) (setq fo (yyc:add-uniq f fo))))))
      (vlax-for lay (vla-get-Layouts dbx)
        (if (and (setq f (vla-get-StyleSheet lay)) (/= f ""))
          (if (setq f (cond ((findfile f)) ((and ctbdir (findfile (strcat ctbdir "\\" f)))))) (setq ct (yyc:add-uniq f ct)))))
      (yyc:dbx-close dbx)))
  (list xr fo ct)
)

;; point every xref in a copied drawing at the file name only (it sits next to it in the zip)
(defun yyc:tx-repath (path / dbx n)
  (setq n 0)
  (if (setq dbx (yyc:dbx-open path))
    (progn
      (vlax-for b (vla-get-Blocks dbx)
        (if (= (vla-get-IsXRef b) :vlax-true)
          (if (not (yyc:err-p (vl-catch-all-apply 'vla-put-Path
                     (list b (strcat (vl-filename-base (vla-get-Path b)) ".dwg")))))
            (setq n (1+ n)))))
      (if (> n 0) (vl-catch-all-apply 'vla-SaveAs (list dbx path)))
      (yyc:dbx-close dbx)))
  n
)

(defun c:YYCTRANSMIT ( / folder dwgno files stage phase groups d code codes plan base zip tmp root deps xr fo ct
                         copied rep f sh cmd ok made pdf fails)
  (setq dwgno (strcase (yyc:get "DwgNo" "24C024")))
  (cond
    ((not (setq folder (yyc:browse-folder "Delivery folder (the finished, renamed DWGs)"))) nil)
    ((not (setq files (vl-remove-if-not '(lambda (x) (wcmatch (strcase x) (strcat dwgno "*"))) (yyc:dwgs-in folder))))
     (yyc:msg (strcat "No drawings starting with " dwgno " in that folder (drawing number set in YYCPROFILES).")))
    (T
     (initget "BID IFC Record")
     (setq stage (getkword (strcat "\nStage [BID/IFC/Record] <" (yyc:get "TxStage" "Record") ">: ")))
     (if (not stage) (setq stage (yyc:get "TxStage" "Record")))
     (setq phase (strcase (yyc:ask "Phase (PH1, PH2... or NONE if the project isn't phased)" (yyc:get "TxPhase" "PH1"))))
     (yyc:set "TxStage" stage) (yyc:set "TxPhase" phase)
     ;; group by discipline letter
     (foreach x files
       (setq d (yyc:fd-disc x dwgno))
       (if (assoc d groups) (setq groups (subst (cons d (append (cdr (assoc d groups)) (list x))) (assoc d groups) groups))
         (setq groups (append groups (list (cons d (list x)))))))
     (foreach g groups
       (setq code (cond ((cdr (assoc (car g) *yyc-zip-codes*))) ((strcat (car g) (car g)))))
       (if (not (member (car g) '("A" "E" "M" "S")))
         (setq code (strcase (yyc:ask (strcat "Two-letter zip code for discipline " (car g)) code))))
       (setq codes (cons (cons (car g) code) codes)
             base (yyc:join (vl-remove "" (list dwgno stage (if (= phase "NONE") "" phase) code)) "_")
             plan (cons (list (car g) base (cdr g)) plan)))
     (setq plan (reverse plan))
     (yyc:msg "Packages:")
     (foreach p plan (yyc:msg (strcat "  " (cadr p) ".zip   " (itoa (length (caddr p))) " drawing(s)")))
     (if (yyc:yes "Make these zips?" "Yes")
       (progn
         (setq root (strcat folder "\\_YYC_transmit") sh (vlax-create-object "WScript.Shell"))
         (vl-mkdir root)
         (foreach p plan
           (setq base (cadr p) tmp (strcat root "\\" base) zip (strcat folder "\\" base ".zip") copied nil rep nil)
           (yyc:msg (strcat "Packing " base ".zip ..."))
           (vl-mkdir tmp)
           (foreach x (caddr p)
             (setq f (strcat folder "\\" x) deps (yyc:tx-deps f))
             (if (findfile (strcat tmp "\\" x)) (vl-file-delete (strcat tmp "\\" x)))
             (if (vl-file-copy f (strcat tmp "\\" x)) (setq copied (yyc:add-uniq x copied)) (setq fails (cons x fails)))
             (setq rep (append rep (list (strcat x (if (car deps) (strcat "  (xrefs: " (yyc:join (mapcar '(lambda (q) (strcat (vl-filename-base q) ".dwg")) (car deps)) ", ") ")") "")))))
             (foreach q (append (car deps) (cadr deps) (caddr deps))
               (if (not (findfile (strcat tmp "\\" (vl-filename-base q) (cond ((vl-filename-extension q)) ("")))))
                 (progn (vl-file-copy q (strcat tmp "\\" (vl-filename-base q) (cond ((vl-filename-extension q)) (""))))
                        (setq copied (yyc:add-uniq (strcat (vl-filename-base q) (cond ((vl-filename-extension q)) (""))) copied)))))
             (yyc:tx-repath (strcat tmp "\\" x)))
           ;; YYC linetype shapes, always
           (foreach q (list (findfile "CAA.SHX") (findfile "CAA.LIN"))
             (if (and q (not (findfile (strcat tmp "\\" (vl-filename-base q) (vl-filename-extension q)))))
               (progn (vl-file-copy q (strcat tmp "\\" (vl-filename-base q) (vl-filename-extension q)))
                      (setq copied (yyc:add-uniq (strcat (vl-filename-base q) (vl-filename-extension q)) copied)))))
           ;; transmittal report
           (if (setq f (open (strcat tmp "\\Transmittal.txt") "w"))
             (progn
               (foreach l (append (list (strcat base ".zip  -  " (yyc:stamp)) ""
                                        "Drawings (xrefs included, not bound):") rep
                                  (list "" "All files in this package:") copied
                                  (list "" "Made by YYCTRANSMIT (YYC CAD Tools). Fonts, CAA.SHX and plot styles included; xrefs re-pathed to this folder."))
                 (write-line l f))
               (close f)))
           (if (findfile zip) (vl-file-delete zip))
           (setq cmd (strcat "powershell -NoProfile -ExecutionPolicy Bypass -Command \"Compress-Archive -Path "
                             (yyc:ps-q (strcat tmp "\\*")) " -DestinationPath " (yyc:ps-q zip) " -Force\""))
           (vl-catch-all-apply 'vlax-invoke-method (list sh 'Run cmd 0 :vlax-true))
           (if (findfile zip) (setq made (cons (strcat base ".zip") made)) (setq fails (cons (strcat base ".zip") fails))))
         ;; tidy the working folder
         (vl-catch-all-apply 'vlax-invoke-method
           (list sh 'Run (strcat "powershell -NoProfile -Command \"Remove-Item -Recurse -Force " (yyc:ps-q root) "\"") 0 :vlax-true))
         (vlax-release-object sh)
         (yyc:msg (strcat "Made " (itoa (length made)) " zip(s) in " folder ": " (yyc:join (reverse made) ", ")))
         (if fails (yyc:msg (strcat "PROBLEM with: " (yyc:join fails ", "))))
         (setq pdf (vl-directory-files folder "*File Description*.pdf" 1))
         (yyc:msg (if pdf (strcat "File Description PDF beside them: " (car pdf))
                      "No File Description PDF in this folder yet - run YYCFILEDESC and save it here."))
         (yyc:msg "Open one zip and check it once against a normal ETRANSMIT before the first real delivery."))))
  )
  (princ)
)

;;; ---------------------------------------------------------------------------
;;; Help window (YYCHELP), YYC menu, and ribbon loader
;;; ---------------------------------------------------------------------------

;; (step  command  what-it-does). AutoCAD's own commands are marked (AutoCAD).
(setq *yyc-help*
 '(("Setup" "YYCSETUP"      "First time: opens the profiles window. Set the Kit folder, click Fill empty files, add a Description, Save + use.")
   ("Setup" "YYCPROFILES"   "Profiles window: pick a profile and change its description, files and settings right in the window (... buttons browse). Add, remove, Save + use.")
   ("2 Name" "YYCRENAME"    "Rename every DWG in a folder to drawing no. + sheet no. Shows a preview first. Do it before opening the files.")
   ("3 Sheet" "YYCPAGESETUP" "Apply the YYC - Titleblock page setup from the template to every layout, giving a true A0, and moves each sheet so its lower-left corner sits on 0,0. If nothing looks different afterwards, type RE (REGEN).")
   ("4 Sheet" "YYCSCHEDULES" "Click a schedule viewport: everything it touches (crossing) moves to paper space at the same size and spot, then it offers to delete the empty viewport. One at a time. Do this before moving the plan.")
   ("4 Sheet" "CHSPACE"     "(AutoCAD) For anything YYCSCHEDULES can't take, e.g. objects crossing a viewport edge.")
   ("5 Grid" "YYCGRIDIN"    "Insert the YYC grid for this area at 0,0,0 as one block. You pick which grid.")
   ("5 Grid" "YYCALIGNREC"  "First sheet: 4 clicks like ALIGN (source 1, destination 1, source 2, destination 2). Moves model space, never scales, saves the move to the profile.")
   ("5 Grid" "YYCALIGNAPPLY" "Every other sheet: apply the saved move with no clicks, and turn the viewports to match. Will not move a drawing twice.")
   ("6 Clean" "YYCGRIDOUT"  "Remove the grid block once the plan is on the grid.")
   ("6 Clean" "YYCCLEAN"    "Delete empty text, purge everything 3 times, audit and fix.")
   ("6 Clean" "YYCFIND0"    "List names containing $0$ left over from binding xrefs.")
   ("7 Viewport" "YYCVPUCS"    "Give each drawing viewport a UCS that matches its turned view (crosshairs and ortho follow the sheet); the sheet itself stays World. The align tools do this too; use it on drawings aligned earlier.")
   ("7 Viewport" "YYCVPLAYERS" "Put each viewport on its own VIEWPORT# no-plot layer. More VIEWPORT# layers than viewports (e.g. after deleting some)? Run it again, then YYCCLEAN.")
   ("7 Viewport" "YYCVPLOCK" "Lock every viewport and list its scale. Only after the scale and rotation are right.")
   ("8 Layers" "YYCLWDEFAULT" "Set every layer's lineweight to Default (fixes the Revit LineWeight009/025/030).")
   ("8 Layers" "LAYTRANS"   "(AutoCAD) Layer Translator - one pass per job, Map Same first, force BYLAYER in Settings.")
   ("8 Layers" "YYCLAYXL"   "Send this drawing's layers to the Layer Mapper workbook in Excel: what's OK, what's automatic, what's remembered, and what you need to choose (pick from a list).")
   ("8 Layers" "YYCLAYSYNC" "Set colour, linetype and lineweight of every standard layer to match the YYC Layer Reference.")
   ("8 Layers" "YYCLAYMAP"  "Apply: merges every layer into the YYC layer you chose in the workbook (plus the automatic A-WALL-6 -> A-WALL ones) and remembers your choices for next time.")
   ("8 Layers" "YYCLAYWALK" "Walk what's left: a window lists every layer that still isn't YYC. Show it isolates one, pick its YYC layer, Map merges it now. Choices are remembered in the workbook.")
   ("8 Layers" "YYCPARK" "Moves the selected objects (e.g. the layer-0 blocks YYCZERO selects) to the left of the drawing, each in its own spot 200000 mm apart, so you can explode and check them. Saved in the drawing.")
   ("8 Layers" "YYCUNPARK" "Moves everything in the parked spots back by exactly the distance it was moved - exploded pieces included.")
   ("Setup" "YYCMAKEREF" "Once per kit: in a new empty drawing, builds the YYC Layer Reference DWG + DWS (every CADD Manual v6.2 layer with its colour, linetype, plot and description) from the standard CSV, saved into the kit folder for LAYTRANS.")
   ("8 Layers" "YYCLAYEXPORT" "Write all layers to YYC_LayerExport.csv, flagged against the YYC Layer Reference.")
   ("8 Layers" "SETBYLAYER" "(AutoCAD) Force colour and linetype back to BYLAYER. Ctrl+A first, include blocks.")
   ("8 Layers" "YYCZERO"    "Show and select what is on 0 / DEFPOINTS. Never moves anything.")
   ("8 Layers" "LAYWALK"    "(AutoCAD) Walk every layer to check its contents. Tick Restore on exit.")
   ("9 QA" "YYCFINAL"       "Lock viewports, make the main layout active, zoom extents.")
   ("9 QA" "YYCQA"          "CADD Manual 5.15 check. Reports only - writes <drawing>_YYC-QA.txt and YYC_QA_Summary.csv.")
   ("10 Send" "YYCFILELIST" "Write filelist.txt for the File Description and flag badly named files.")
   ("10 Send" "YYCFILEDESC" "Combines every consultant's file list (their .txt / .docx / .csv, or a folder of their DWGs) into one File Description: your last one as the template, today's date, grouped A / S / M / E. Saves Word + PDF.")
   ("10 Send" "YYCTRANSMIT" "One zip per discipline, named <drawing no>_<stage>_<phase>_<AR/ST/ME/EL>.zip. Asks the stage and phase. Includes xrefs (re-pathed, not bound), fonts, CAA.SHX and the CTB, plus a Transmittal.txt.")
   ("10 Send" "ETRANSMIT"   "(AutoCAD) One zip per discipline. Fonts and CTB on, bind xrefs OFF.")
   ("Any" "YYCBATCH"        "Run tools over a whole folder: backs up, saves as AutoCAD 2018, logs each file.")))

(defun yyc:wrap (text width / words line out)
  (setq words (yyc:split text " ") line "")
  (foreach w words
    (if (> (+ (strlen line) (strlen w) 1) width)
      (setq out (cons line out) line w)
      (setq line (if (= line "") w (strcat line " " w)))))
  (reverse (cons line out))
)

(defun yyc:help-show (idx / item lines i)
  (setq item (nth (atoi idx) *yyc-help*) lines (yyc:wrap (caddr item) 92) i 0)
  (set_tile "cmd" (strcat (cadr item) "   -   step " (car item)))
  (foreach k '("d1" "d2" "d3")
    (set_tile k (if (nth i lines) (nth i lines) "")) (setq i (1+ i)))
)

(defun c:YYCHELP ( / dcl f id pick res)
  (setq dcl (vl-filename-mktemp "yychelp.dcl") f (open dcl "w"))
  (foreach x (list
    "yychelp : dialog {"
    (strcat "  label = \"YYC CAD Tools v" *yyc-version* "  -  profile: " (vl-string-translate "\"" "'" (yyc:profile)) "\";")
    "  : text { value = \"Commands in process order. Pick one to read what it does, then Run.\"; }"
    "  : list_box { key = \"list\"; width = 70; height = 22; tabs = \"12\"; }"
    "  : boxed_column { label = \"What it does\";"
    "    : text { key = \"cmd\"; width = 96; }"
    "    : text { key = \"d1\"; width = 96; }"
    "    : text { key = \"d2\"; width = 96; }"
    "    : text { key = \"d3\"; width = 96; }"
    "  }"
    "  : row { : button { key = \"run\"; label = \"Run\"; is_default = true; width = 12; }"
    "          : button { key = \"cancel\"; label = \"Close\"; is_cancel = true; width = 12; } }"
    "}") (write-line x f))
  (close f)
  (setq id (load_dialog dcl))
  (if (and (>= id 0) (new_dialog "yychelp" id))
    (progn
      (start_list "list")
      (foreach h *yyc-help* (add_list (strcat (car h) "\t" (cadr h))))
      (end_list)
      (setq pick "0")
      (set_tile "list" pick)
      (yyc:help-show pick)
      (action_tile "list" "(setq pick $value) (yyc:help-show $value) (if (= $reason 4) (done_dialog 1))")
      (action_tile "run" "(setq pick (get_tile \"list\")) (done_dialog 1)")
      (setq res (start_dialog))
      (unload_dialog id)
      (vl-file-delete dcl)
      (if (and (= res 1) pick (/= pick ""))
        (vla-SendCommand (yyc:doc) (strcat (cadr (nth (atoi pick) *yyc-help*)) " ")))
    )
    (yyc:msg "Could not open the help window.")
  )
  (princ)
)

;; Earlier versions added a "YYC" drop-down to AutoCAD's main menu at load. Editing
;; the main menu while AutoCAD runs can upset its interface (palettes such as Layer
;; Properties going missing), so that's gone. This removes the menu if an earlier
;; version left one behind. Help is YYCHELP; the ribbon tab is a separate YYC.cuix.
(defun yyc:remove-old-menu ( / mg pm)
  (vlax-for g (vla-get-MenuGroups (yyc:acad))
    (if (setq pm (yyc:try 'vla-Item (list (vla-get-Menus g) "YYC")))
      (if (= (vla-get-OnMenuBar pm) :vlax-true) (vl-catch-all-apply 'vla-RemoveFromMenuBar (list pm)))))
)

;; folder this file was loaded from (bundle or plain folder)
(defun yyc:home ( / p)
  (setq p (cond ((findfile "YYC-CADTools.lsp"))
                ((findfile (strcat (getenv "APPDATA") "\\Autodesk\\ApplicationPlugins\\YYC-CADTools.bundle\\Contents\\YYC-CADTools.lsp")))
                ((findfile (strcat (getenv "PROGRAMDATA") "\\Autodesk\\ApplicationPlugins\\YYC-CADTools.bundle\\Contents\\YYC-CADTools.lsp")))
                ((yyc:get "ToolsPath" nil))))
  (if p (vl-filename-directory p))
)

;; load the YYC ribbon tab (YYC.cuix) if one sits next to this file
(defun yyc:load-ribbon ( / home cuix mgs loaded)
  (setq home (yyc:home) mgs (vla-get-MenuGroups (yyc:acad)))
  (if (and home (setq cuix (findfile (strcat home "\\YYC.cuix"))))
    (progn
      (vlax-for g mgs (if (= (strcase (vla-get-Name g)) "YYC") (setq loaded T)))
      (if (not loaded) (vl-catch-all-apply 'vla-Load (list mgs cuix)))
    )
  )
)

;;; ---------------------------------------------------------------------------
;;; Batch - opens each drawing through ActiveX (Documents.Open), runs the tools
;;; on that document object, saves as AutoCAD 2018 and closes. No script file.
;;; ---------------------------------------------------------------------------

(setq *yyc-batch-cmds*
  '(("Align" . yyc:t-alignapply) ("Pagesetup" . yyc:t-pagesetup) ("Vplayers" . yyc:t-vplayers) ("Lwdefault" . yyc:t-lwdefault) ("Laysync" . yyc:t-laysync)
    ("Mapcsv" . yyc:t-laymap) ("Clean" . yyc:t-clean) ("Dollar0" . yyc:t-find0)
    ("Export" . yyc:t-layexport) ("Qa" . yyc:t-qa) ("Final" . yyc:t-final)))
(setq *yyc-readonly* '(yyc:t-find0 yyc:t-layexport yyc:t-qa))

(defun yyc:log (folder line / f)
  (if (setq f (open (strcat folder "\\YYC_Batch_Log.txt") "a"))
    (progn (write-line (strcat (yyc:stamp) "  " line) f) (close f)))
)

(defun yyc:open-docs ( / res)
  (vlax-for d (vla-get-Documents (yyc:acad)) (setq res (cons (strcase (vla-get-FullName d)) res)))
  res
)

(defun c:YYCBATCH ( / *error* folder files opn pick seq kw bk save fmt docs d path r ok n)
  (defun *error* (m) (setq *yyc-batch* nil *yyc-cur* nil) (if m (princ (strcat "\nBatch stopped: " m))) (princ))
  (cond
    ((/= (getvar "SDI") 0) (yyc:msg "YYCBATCH needs SDI = 0 (several drawings open at once)."))
    ((not (setq folder (yyc:browse-folder "Folder of drawings to process"))) (yyc:msg "Cancelled."))
    (T
     (setq files (yyc:dwgs-in folder) opn (yyc:open-docs))
     (yyc:msg (strcat (itoa (length files)) " drawing(s) in " folder))
     (yyc:msg "Pick tools in the order to run them. Enter when done.")
     (setq kw (strcat (yyc:join (mapcar 'car *yyc-batch-cmds*) " ") " Run"))
     (while (progn (initget kw)
                   (setq pick (getkword (strcat "\nAdd [" (vl-string-translate " " "/" kw) "] <Run>: ")))
                   (and pick (/= pick "Run")))
       (setq seq (append seq (list (cdr (assoc pick *yyc-batch-cmds*)))))
       (yyc:msg (strcat "  Sequence: " (yyc:join (mapcar '(lambda (s) (car (vl-some '(lambda (p) (if (eq (cdr p) s) p)) *yyc-batch-cmds*))) seq) " > ")))
     )
     (cond
       ((not seq) (yyc:msg "Nothing picked."))
       ;; ask for every file the tools need now, so nothing stops mid-batch
       ((and (member 'yyc:t-pagesetup seq) (not (yyc:need-file "Template" "Select the YYC titleblock template" "dwt"))) (yyc:msg "Cancelled."))
       ((and (member 'yyc:t-laymap seq) (not (yyc:need-file "LayerMap" "Select the layer map CSV" "csv"))) (yyc:msg "Cancelled."))
       (T
        (setq save (vl-some '(lambda (s) (not (member s *yyc-readonly*))) seq))
        (if (and save (yyc:yes "Back up the drawings first?" "Yes"))
          (progn
            (setq bk (strcat folder "\\_YYC_backup_" (yyc:stamp-file)))
            (vl-mkdir bk)
            (foreach x files (vl-file-copy (strcat folder "\\" x) (strcat bk "\\" x)))
            (yyc:msg (strcat "Backed up to " bk))
          )
        )
        (setq fmt (cond ((eval 'ac2018_dwg)) (64)) docs (vla-get-Documents (yyc:acad)) n 0 *yyc-batch* T)
        (yyc:log folder (strcat "---- batch start, profile " (yyc:profile) ", " (itoa (length files)) " file(s), saving: " (if save "yes, AutoCAD 2018" "no (read-only tools)")))
        (foreach x files
          (setq path (strcat folder "\\" x) *yyc-cur* nil)
          (cond
            ((member (strcase path) opn) (yyc:msg (strcat "Skipped (open right now): " x)) (yyc:log folder (strcat x "  SKIPPED - open")))
            ((yyc:err-p (setq d (vl-catch-all-apply 'vla-Open (list docs path :vlax-false))))
             (yyc:msg (strcat "Could not open " x)) (yyc:log folder (strcat x "  FAILED to open: " (vl-catch-all-error-message d))))
            (T
             (setq *yyc-cur* x ok T)
             (foreach fn seq
               (if (yyc:err-p (setq r (vl-catch-all-apply fn (list d))))
                 (progn (setq ok nil) (yyc:msg (strcat "ERROR in " (vl-symbol-name fn) ": " (vl-catch-all-error-message r))))))
             (if save
               (if (yyc:err-p (setq r (vl-catch-all-apply 'vla-SaveAs (list d path fmt))))
                 (setq ok nil r (strcat "SAVE FAILED: " (vl-catch-all-error-message r)))
                 (setq r "saved as AutoCAD 2018"))
               (setq r "not saved (read-only tools)"))
             (vl-catch-all-apply 'vla-Close (list d :vlax-false))
             (setq n (1+ n))
             (yyc:log folder (strcat x "  " (if ok "OK" "CHECK") "  " r))
            )
          )
        )
        (setq *yyc-batch* nil *yyc-cur* nil)
        (yyc:log folder "---- batch end")
        (yyc:msg (strcat "Batch finished: " (itoa n) " drawing(s). See YYC_Batch_Log.txt in the folder."))
       )
     )
    )
  )
  (princ)
)

(setq *yyc-batch* nil *yyc-cur* nil)
(if (not *yyc-menu-cleaned*) (progn (vl-catch-all-apply 'yyc:remove-old-menu nil) (setq *yyc-menu-cleaned* T)))
(vl-catch-all-apply 'yyc:load-ribbon nil)
;; put the profile's kit folder on AutoCAD's support path, so CAA.SHX and the
;; caa_eng / caa_arch fonts kept there are found (needed to show YYC linetypes and text)
(defun yyc:ensure-support ( / kit files sp)
  (if (and (setq kit (yyc:get "Kit" nil)) (vl-file-directory-p kit)
           (setq files (yyc:try 'vla-get-Files (list (vla-get-Preferences (yyc:acad)))))
           (setq sp (yyc:try 'vla-get-SupportPath (list files)))
           (not (member (strcase kit) (mapcar 'strcase (yyc:split sp ";")))))
    (if (not (yyc:err-p (vl-catch-all-apply 'vla-put-SupportPath (list files (strcat sp ";" kit)))))
      (princ (strcat "\nYYC: kit folder added to the support path (for CAA.SHX and the CAA fonts)."))))
)
(yyc:try 'yyc:ensure-support nil)

(princ (strcat "\nYYC CAD Tools v" *yyc-version* " (ActiveX) loaded. Type YYCHELP for every command in order. Profile: " (yyc:profile) ". Commands: YYCHELP YYCPROFILES YYCRENAME YYCPAGESETUP YYCSCHEDULES YYCGRIDIN YYCALIGNREC YYCALIGNAPPLY YYCVPUCS YYCGRIDOUT YYCCLEAN YYCFIND0 YYCVPLAYERS YYCVPLOCK YYCLWDEFAULT YYCLAYXL YYCLAYMAP YYCLAYWALK YYCLAYSYNC YYCMAKEREF YYCLAYEXPORT YYCZERO YYCPARK YYCUNPARK YYCQA YYCFINAL YYCFILELIST YYCFILEDESC YYCTRANSMIT YYCBATCH"))
(princ)
