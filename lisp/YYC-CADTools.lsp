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
;;;   Step 3   YYCTITLEBLOCK  (experimental) Swap the exported titleblock for the YYC A0 one
;;;   Step 3   YYCPAGESETUP   Copy "YYC - Titleblock" page setup from the .dwt to all layouts
;;;   Step 4   YYCSCHEDULES   Pick a viewport: its contents go to paper space, then delete it
;;;   Step 5   YYCGRIDIN      Insert the YYC grid (04grid-dtb.dwg) as a block at 0,0,0
;;;            YYCALIGNREC    Do the MOVE/ALIGN once with 4 clicks; saves it to the profile
;;;            YYCALIGNAPPLY  Apply the saved move to this drawing (or Align in YYCBATCH)
;;;            YYCVPFOLLOW    Turn viewports to match the move (automatic in the align tools)
;;;            YYCALIGNUSE    Let this profile use another profile's saved move
;;;   Step 6   YYCGRIDOUT     Delete the grid block and its definition
;;;            YYCCLEAN       Delete empty text, PurgeAll x3, AuditInfo (fix)
;;;            YYCFIND0       Report named objects containing "$0$" (bind leftovers)
;;;   Step 7   YYCVPLAYERS    Put every viewport on its own VIEWPORT# no-plot layer
;;;            YYCVPLOCK      Lock every viewport and list its scale
;;;   Step 8   YYCLWDEFAULT   Set every layer's lineweight to Default
;;;            YYCLAYMAP      Merge layers using a CSV map (old,new)
;;;            YYCLAYEXPORT   Write layers to YYC_LayerExport.csv, checked against the
;;;                           YYC Layer Reference
;;;            YYCZERO        Report and select objects on 0 / DEFPOINTS (report only)
;;;   Step 9   YYCQA          CADD Manual 5.15 check - report only, changes nothing
;;;            YYCFINAL       Lock viewports, main layout active, zoom extents
;;;   Step 10  YYCFILELIST    Write filelist.txt for a folder and check file names
;;;   Any      YYCBATCH       Run a chosen sequence of the above over a folder
;;;            YYCSETUP (first time) / YYCPROFILES (everything else about profiles)
;;;   Help     YYCHELP        Every command in process order, with a Run button
;;;            A "YYC" menu is added to the menu bar (MENUBAR 1 to show it), and a
;;;            YYC.cuix ribbon tab next to this file is loaded automatically.
;;;
;;; Profiles: one set of kit files per project or area (grid, template, layer
;;; reference...). YYCPROFILES switches; YYCGRIDIN lets you pick the grid each time.
;;; ===========================================================================

(vl-load-com)
(setq *yyc-version* "0.8.1")
(setq *yyc-modifiers* '("DEMO" "EXST" "FUTR" "MOVE" "NEWW" "NICN" "NPLT" "PRPS" "RELO" "TEMP"))

;;; ---------------------------------------------------------------------------
;;; Core helpers
;;; ---------------------------------------------------------------------------

(defun yyc:acad () (vlax-get-acad-object))
(defun yyc:doc () (vla-get-ActiveDocument (yyc:acad)))

;; Settings live in named PROFILES (e.g. "24C024 DTB", "ITB"), one set of kit
;; files per project or area. Profile / Profiles / ToolsPath are shared.
(setq *yyc-keys* '("Description" "Kit" "Template" "Titleblock" "PageSetup" "Grid" "GridBlock" "LayerRef" "LinFile" "LayerMap" "DwgNo" "CTB" "Fonts"))

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

(defun yyc:browse-folder (msg / sh fo path)
  (setq sh (vlax-create-object "Shell.Application"))
  (setq fo (vl-catch-all-apply 'vlax-invoke-method (list sh 'BrowseForFolder 0 msg 0)))
  (if (and fo (not (yyc:err-p fo)))
    (progn
      (setq path (yyc:try '(lambda () (vlax-get-property (vlax-get-property fo 'Self) 'Path)) nil))
      (vlax-release-object fo)
    )
  )
  (vlax-release-object sh)
  (if (and path (vl-file-directory-p path)) path nil)
)

(defun yyc:dwgs-in (folder / files)
  (setq files (vl-directory-files folder "*.dwg" 1))
  (if files (vl-sort files '(lambda (a b) (< (strcase a) (strcase b)))) nil)
)

;; every kind of kit file the tools use: key, label, search patterns, extension
(setq *yyc-file-kinds*
  '(("Template" "Titleblock template (.dwt)" ("*.dwt") "dwt")
    ("Grid"     "Grid drawing (04 DTB domestic, 20 ITB international ...)" ("*grid*.dwg") "dwg")
    ("Titleblock" "YYC A0 titleblock drawing" ("*titleblock*.dwg") "dwg")
    ("LayerRef" "YYC Layer Reference" ("*layer*.dwg" "*.dws") "dwg")
    ("LinFile"  "CAA linetype file" ("*.lin") "lin")
    ("LayerMap" "Layer map CSV (old,new)" ("*.csv") "csv")))

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
    (setq pick (getfiled (nth 1 kind) (yyc:get "Kit" "") (nth 3 kind) 0))
    (progn
      (yyc:msg (strcat (nth 1 kind) " - profile \"" (yyc:profile) "\":"))
      (setq i 0)
      (foreach c cands
        (setq i (1+ i))
        (yyc:msg (strcat (if (and cur (= (strcase c) (strcase cur))) "  * " "    ") (itoa i) ". "
                         (vl-filename-base c) (vl-filename-extension c) "   (" (vl-filename-directory c) ")")))
      (setq ans (getstring (strcat "\nNumber, B to browse" (if cur (strcat ", Enter keeps " (vl-filename-base cur)) "") ": ")))
      (cond
        ((= ans "") (setq pick cur))
        ((= (strcase ans) "B") (setq pick (getfiled (nth 1 kind) (yyc:get "Kit" "") (nth 3 kind) 0)))
        ((and (> (atoi ans) 0) (<= (atoi ans) (length cands))) (setq pick (nth (1- (atoi ans)) cands)))
        (T (yyc:msg "Not a choice - kept the current one.") (setq pick cur))
      )
    )
  )
  (if pick (yyc:set key pick))
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

(defun c:YYCSETUP ( / name kit)
  (yyc:msg (strcat "YYC CAD Tools setup. Profiles: " (yyc:join (yyc:profiles) ", ")))
  (yyc:msg "A profile is one set of kit files - make one per project or area (e.g. \"24C024 DTB\", \"ITB\").")
  (setq name (yyc:ask "Profile to set up" (yyc:profile)))
  (if (/= (strcase name) (strcase (yyc:profile)))
    (yyc:switch-profile name (and (not (member (strcase name) (mapcar 'strcase (yyc:profiles))))
                                  (yyc:yes (strcat "Start \"" name "\" from a copy of \"" (yyc:profile) "\"?") "Yes")))
    (yyc:add-profile name)
  )
  (yyc:set "Description" (yyc:ask "Short description (e.g. 24C024 DTB departures, Area 1-3)" (yyc:get "Description" "")))
  (yyc:msg (strcat "Kit folder now: " (yyc:get "Kit" "(not set)") " - pick a new one, or Cancel to keep it."))
  (if (setq kit (yyc:browse-folder (strcat "Kit folder for profile \"" (yyc:profile) "\" (Cancel keeps the current one)")))
    (yyc:set "Kit" kit))
  (foreach kind *yyc-file-kinds* (yyc:choose-file (car kind)))
  (yyc:set "PageSetup" (yyc:ask "Page setup name inside the template" (yyc:get "PageSetup" "YYC - Titleblock")))
  (yyc:set "DwgNo"     (yyc:ask "YYC drawing number" (yyc:get "DwgNo" "24C024")))
  (yyc:set "CTB"       (yyc:ask "Plot style table" (yyc:get "CTB" "YYC_BW_HPv1.ctb")))
  (yyc:set "Fonts"     (yyc:ask "Allowed font files, comma separated" (yyc:get "Fonts" "caa_eng.shx,caa_arch.shx,CAA.SHX")))
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

(defun c:YYCRENAME ( / folder dwgno files pairs new ok fail)
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
          (yyc:msg "---- Preview (old  ->  new) ----")
          (foreach p pairs (yyc:msg (strcat "  " (car p) "  ->  " (cdr p))))
          (yyc:msg "Close these drawings first. If the Revit export wrote views as xrefs, renaming those xref files breaks their paths.")
          (if (yyc:yes (strcat "Rename " (itoa (length pairs)) " file(s)?") "No")
            (progn
              (setq ok 0 fail 0)
              (foreach p pairs
                (if (and (not (findfile (strcat folder "\\" (cdr p))))
                         (vl-file-rename (strcat folder "\\" (car p)) (strcat folder "\\" (cdr p))))
                  (setq ok (1+ ok))
                  (progn (setq fail (1+ fail)) (yyc:msg (strcat "  NOT renamed (open, or name already taken): " (car p))))
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
              (progn (setq n (1+ n)) (yyc:try 'vla-RefreshPlotDeviceInfo (list lay))
                     (yyc:msg (strcat "  " (vla-get-Name lay) ": " (vla-get-CanonicalMediaName lay) ", " (vla-get-StyleSheet lay))))
            )
          )
          (yyc:msg (strcat "Page setup \"" ps "\" applied to " (itoa n) " layout(s)."))
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
        (yyc:msg "Next: rough MOVE, then ALIGN with two distant points, answer No to scale. Then YYCGRIDOUT.")
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
;;; Titleblock swap (experimental) - replace the Revit-exported titleblock with
;;; the YYC A0 titleblock and carry the exported text into its attributes.
;;; The YYC titleblock drawing is read with ObjectDBX: everything in its first
;;; layout (inside the sheet) becomes block YYC-A0-TITLEBLOCK, inserted at 0,0.
;;; ---------------------------------------------------------------------------

(setq *yyc-tb-fields* '(1060.0 0.0 1175.0 160.0))   ; where the exported title fields sit
(setq *yyc-tb-revs*   '(1065.0 600.0 1190.0 672.0)) ; exported revision table
(setq *yyc-tb-sheet*  '(-20.0 -20.0 1210.0 870.0))  ; anything outside is not copied

(defun yyc:in-box (x y box) (and (>= x (nth 0 box)) (>= y (nth 1 box)) (<= x (nth 2 box)) (<= y (nth 3 box))))
;; whole object inside the box (objects without extents count as inside)
(defun yyc:obj-in-box (o box / lo hi)
  (if (yyc:err-p (vl-catch-all-apply 'vla-GetBoundingBox (list o 'lo 'hi)))
    T
    (progn (setq lo (vlax-safearray->list lo) hi (vlax-safearray->list hi))
           (and (yyc:in-box (car lo) (cadr lo) box) (yyc:in-box (car hi) (cadr hi) box))))
)
(defun yyc:pt (o) (vlax-safearray->list (vlax-variant-value (vla-get-InsertionPoint o))))

;; plain text out of MTEXT formatting codes
(defun yyc:plain (s / out i ch nx)
  (setq out "" i 1)
  (while (<= i (strlen s))
    (setq ch (substr s i 1))
    (cond
      ((member ch '("{" "}")) (setq i (1+ i)))
      ((= ch "\\")
       (setq nx (substr s (1+ i) 1))
       (cond
         ((member nx '("P" "~")) (setq out (strcat out " ") i (+ i 2)))
         ((member nx '("\\" "{" "}")) (setq out (strcat out nx) i (+ i 2)))
         ((member nx '("L" "l" "O" "o" "K" "k")) (setq i (+ i 2)))
         (T (setq i (+ i 2)) (while (and (<= i (strlen s)) (/= (substr s i 1) ";")) (setq i (1+ i))) (setq i (1+ i))))
      )
      (T (setq out (strcat out ch) i (1+ i)))
    )
  )
  (while (vl-string-search "  " out) (setq out (vl-string-subst " " "  " out)))
  (yyc:trim out)
)

(defun yyc:tb-def (doc / name blk path dbx lay objs arr r)
  (setq name "YYC-A0-TITLEBLOCK")
  (cond
    ((setq blk (yyc:try 'vla-Item (list (vla-get-Blocks doc) name))) blk)
    ((not (setq path (yyc:need-file "Titleblock" "Select the YYC A0 titleblock drawing" "dwg"))) nil)
    ((not (setq dbx (yyc:dbx-open path))) nil)
    (T
     (setq lay (car (yyc:paper-layouts dbx)))
     (vlax-for o (vla-get-Block lay)
       (if (and (/= (vla-get-ObjectName o) "AcDbViewport") (yyc:obj-in-box o *yyc-tb-sheet*))
         (setq objs (cons o objs))))
     (setq blk (vla-Add (vla-get-Blocks doc) (vlax-3d-point 0 0 0) name))
     (setq arr (vlax-make-safearray vlax-vbObject (cons 0 (1- (length objs)))))
     (vlax-safearray-fill arr objs)
     (setq r (vl-catch-all-apply 'vla-CopyObjects (list dbx arr blk)))
     (if (yyc:err-p r)
       (progn   ; retry without OLE objects (logos) if they refuse to copy
         (setq objs (vl-remove-if '(lambda (o) (= (vla-get-ObjectName o) "AcDbOle2Frame")) objs))
         (setq arr (vlax-make-safearray vlax-vbObject (cons 0 (1- (length objs)))))
         (vlax-safearray-fill arr objs)
         (setq r (vl-catch-all-apply 'vla-CopyObjects (list dbx arr blk)))
         (if (not (yyc:err-p r)) (yyc:msg "Note: the OLE logo could not be copied - paste it by hand once."))))
     (yyc:dbx-close dbx)
     (if (yyc:err-p r) (progn (yyc:msg (strcat "ERROR copying the titleblock: " (vl-catch-all-error-message r))) nil) blk))
  )
)

;; exported title text in one layout, merged into lines: ((x y "text" objs) ...)
(defun yyc:tb-lines (lay / items lines cur)
  (vlax-for o (vla-get-Block lay)
    (if (and (yyc:text-p o) (= (strcase (vla-get-Layer o)) (strcase (yyc:get "TBLayer" "G-ANNO-TTLB"))))
      (setq items (cons (list (car (yyc:pt o)) (cadr (yyc:pt o)) (yyc:plain (vla-get-TextString o)) o) items))))
  (setq items (vl-remove-if-not '(lambda (i) (yyc:in-box (car i) (cadr i) *yyc-tb-fields*)) items))
  (setq items (vl-sort items '(lambda (a b) (if (> (abs (- (cadr a) (cadr b))) 2.5) (> (cadr a) (cadr b)) (< (car a) (car b))))))
  (foreach i items
    (if (and cur (<= (abs (- (cadr i) (cadr cur))) 2.5))
      (setq cur (list (car cur) (cadr cur) (yyc:trim (strcat (caddr cur) " " (caddr i))) (cons (cadddr i) (nth 3 cur))))
      (progn (if cur (setq lines (cons cur lines))) (setq cur (list (car i) (cadr i) (caddr i) (list (cadddr i))))))
  )
  (if cur (setq lines (cons cur lines)))
  (vl-remove-if '(lambda (l) (= (caddr l) "")) (reverse lines))
)

;; match exported lines to attributes: nearest first (rows count more than columns),
;; then numbered groups (PROJ-NAME-1..3, SHEET-TITLE-1..3) take their lines top-down
(defun yyc:tb-match (atts lines / pairs used-a used-l res base mem ys cands)
  (foreach a atts
    (foreach l lines
      (setq pairs (cons (list (+ (* 0.25 (abs (- (cadr a) (car l)))) (abs (- (caddr a) (cadr l)))) a l) pairs))))
  (foreach p (vl-sort pairs '(lambda (x y) (< (car x) (car y))))
    (if (and (< (car p) 9.0) (not (member (cadr p) used-a)) (not (member (caddr p) used-l)))
      (setq used-a (cons (cadr p) used-a) used-l (cons (caddr p) used-l) res (cons (cons (cadr p) (caddr p)) res))))
  (foreach base '("PROJ-NAME-" "SHEET-TITLE-")
    (setq mem (vl-sort (vl-remove-if-not '(lambda (a) (wcmatch (strcase (car a)) (strcat base "#"))) atts)
                       '(lambda (x y) (< (car x) (car y)))))
    (if mem
      (progn
        (setq ys (mapcar 'caddr mem))
        (setq cands (vl-remove-if-not
                      '(lambda (l) (and (<= (cadr l) (+ (apply 'max ys) 8)) (>= (cadr l) (- (apply 'min ys) 8))
                                        (<= (abs (- (car l) (cadr (car mem)))) 40)))
                      lines))
        (setq cands (vl-sort cands '(lambda (x y) (> (cadr x) (cadr y)))))
        (setq res (vl-remove-if '(lambda (r) (or (member (car r) mem) (member (cdr r) cands))) res))
        (setq ys mem)
        (foreach l cands (if ys (setq res (cons (cons (car ys) l) res) ys (cdr ys))))
      )
    )
  )
  res
)

(defun yyc:t-titleblock (doc / blk n ref atts lines m old revs main sc dn prompts)
  (if (setq blk (yyc:tb-def doc))
    (progn
      (setq n 0)
      (foreach pr (yyc:vp-list doc)   ; main viewport = the biggest one
        (if (or (not main) (> (* (vla-get-Width (cdr pr)) (vla-get-Height (cdr pr))) (* (vla-get-Width (cdr main)) (vla-get-Height (cdr main)))))
          (setq main pr)))
      (vlax-for o blk
        (if (= (vla-get-ObjectName o) "AcDbAttributeDefinition")
          (setq prompts (cons (cons (strcase (vla-get-TagString o)) (strcase (vla-get-PromptString o))) prompts))))
      (foreach lay (yyc:paper-layouts doc)
        (setq old nil revs nil)
        (vlax-for o (vla-get-Block lay)
          (cond
            ((and (= (vla-get-ObjectName o) "AcDbBlockReference") (wcmatch (strcase (vla-get-Name o)) "*TITLEBLOCK*")
                  (/= (strcase (vla-get-Name o)) "YYC-A0-TITLEBLOCK"))
             (setq old (cons o old)))
            ((and (yyc:text-p o) (apply 'yyc:in-box (append (yyc:first-n (yyc:pt o) 2) (list *yyc-tb-revs*))))
             (setq revs (cons o revs)))))
        (if (not old)
          (yyc:msg (strcat "  " (vla-get-Name lay) ": no exported titleblock found - skipped."))
          (progn
            (setq lines (yyc:tb-lines lay))
            (setq ref (vla-InsertBlock (vla-get-Block lay) (vlax-3d-point 0 0 0) "YYC-A0-TITLEBLOCK" 1.0 1.0 1.0 0.0))
            (yyc:try 'vla-put-Layer (list ref "0"))
            (setq atts (mapcar '(lambda (a) (list (vla-get-TagString a) (car (yyc:pt a)) (cadr (yyc:pt a)) a))
                               (if (= (vla-get-HasAttributes ref) :vlax-true) (vlax-invoke ref 'GetAttributes))))
            (yyc:msg (strcat "  " (vla-get-Name lay) ":"))
            (foreach m (yyc:tb-match atts lines)
              (vla-put-TextString (nth 3 (car m)) (caddr (cdr m)))
              (foreach t2 (nth 3 (cdr m)) (vl-catch-all-apply 'vla-Delete (list t2)))
              (setq lines (vl-remove (cdr m) lines))
              (yyc:msg (strcat "    " (car (car m)) "  <-  " (caddr (cdr m))))
            )
            ;; fields the export doesn't carry
            (foreach a atts
              (cond
                ((= (cdr (assoc (strcase (car a)) prompts)) "ENTER CADD FILE NUMBER")
                 (setq dn (yyc:dbase doc))
                 (vla-put-TextString (nth 3 a) (if (wcmatch (strcase dn) (strcat (strcase (yyc:get "DwgNo" "")) "*")) dn (yyc:get "DwgNo" "-")))
                 (yyc:msg (strcat "    " (car a) "  <-  " (vla-get-TextString (nth 3 a)) "  (drawing file name)")))
                ((and (= (strcase (car a)) "SCALE") main)
                 (vla-put-TextString (nth 3 a) (yyc:scale-text (cdr main)))
                 (yyc:msg (strcat "    SCALE  <-  " (yyc:scale-text (cdr main)) "  (main viewport)")))))
            (foreach l lines (yyc:msg (strcat "    left as exported text (no matching field): " (caddr l))))
            ;; revision table: drop the exported headings (YYC's are in the titleblock), keep the rows plottable
            (foreach r revs
              (if (member (strcase (yyc:plain (vla-get-TextString r))) '("NO" "NO." "DATE" "ISSUED FOR" "BY"))
                (vl-catch-all-apply 'vla-Delete (list r))
                (yyc:try 'vla-put-Layer (list r "T-TTLB-REVS"))))
            (if revs (yyc:msg "    revision rows moved to T-TTLB-REVS (they were exported on a no-plot layer)"))
            (foreach o old (vl-catch-all-apply 'vla-Delete (list o)))
            (setq n (1+ n))
          )
        )
      )
      (yyc:msg (strcat "YYC A0 titleblock placed on " (itoa n) " layout(s). Check each field; edit any with ATTEDIT (double-click the titleblock)."))
      (yyc:msg "Then YYCCLEAN to purge the old exported titleblock block.")
    )
  )
)
(defun c:YYCTITLEBLOCK () (yyc:run 'yyc:t-titleblock))

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

;; Model objects currently visible in the viewport: on a visible layer and inside
;; the viewport window, allowing a small margin past the frame (MARGIN paper mm)
;; for border lines that sit right under the frame edge.
;; Returns (visible locked-count partly-count)
(defun yyc:vp-contents (doc e g margin / fp w h vc tg tw k hw hh lo hi in part lock pts ok any hid lay)
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
        (setq ok T any nil)
        (foreach q pts
          (if (and (<= (abs (- (car q) (car vc))) hw) (<= (abs (- (cadr q) (cadr vc))) hh))
            (setq any T) (setq ok nil)))
        (cond
          ((and ok (= (vla-get-Lock (vla-Item (vla-get-Layers doc) (vla-get-Layer o))) :vlax-true)) (setq lock (1+ lock)))
          (ok (setq in (cons o in)))
          (any (setq part (1+ part))))
      )
    )
  )
  (list in lock part)
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
  (yyc:msg (strcat "Only what's visible in the viewport (layers on, not frozen there), plus " (yyc:get "SchedMargin" "5") " mm past the frame for border lines."))
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
               (yyc:msg (strcat "Found " (itoa (length objs)) " object(s) fully inside: " (yyc:type-summary objs)))
               (if (> (caddr res) 0)
                 (yyc:msg (strcat "  " (itoa (caddr res)) " more run past the edge (beyond the "
                                  (yyc:get "SchedMargin" "5") " mm allowance) - those stay in model space.")))
               (if (> (cadr res) 0)
                 (yyc:msg (strcat "  " (itoa (cadr res)) " are on locked layers - skipped. Unlock them and run again to include them.")))
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
                         (progn (vl-catch-all-apply 'vla-Delete (list o)) (yyc:msg "Viewport deleted."))
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
      (setq n (1+ n))
      (yyc:msg (strcat "  Could not turn the viewport in " (car pr))))
    (vl-catch-all-apply 'vla-put-DisplayLocked (list vp lk))
  )
  (yyc:set-mark doc "YYC_VPFollowed" (yyc:stamp))
  (yyc:msg (strcat (itoa n) " viewport(s) now show the same area as the export, turned to project north (twist "
                   (rtos (yyc:deg (- th)) 2 3) " deg). Check one; if it is rotated the wrong way, tell us."))
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

(defun yyc:vp-follow-run (doc ask pts / vps all)
  (setq vps (yyc:vp-choice doc ask) all (yyc:vp-list doc))
  (if vps (apply 'yyc:vp-follow (append (list doc vps) pts)))
  (foreach pr all
    (if (not (member (vla-get-Handle (cdr pr)) (mapcar '(lambda (x) (vla-get-Handle (cdr x))) vps)))
      (yyc:msg (strcat "  Left as it was: viewport in " (car pr) " (" (rtos (vla-get-Width (cdr pr)) 2 0) " x "
                       (rtos (vla-get-Height (cdr pr)) 2 0) ") - it still looks at the old location. Delete it if it's empty."))))
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
(defun c:YYCVPFOLLOW () (yyc:run 'yyc:t-vpfollow))

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
    ((member nil pts) (yyc:msg (strcat "No recorded move for profile \"" (yyc:profile) "\". Run YYCALIGNREC, or YYCALIGNUSE to borrow one.")))
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
(defun c:YYCALIGNUSE ( / ps i ans)
  (setq ps (vl-remove-if-not '(lambda (x) (yyc:get-in x "AlignS1" nil)) (yyc:profiles)) i 0)
  (if (not ps)
    (yyc:msg "No profile has a recorded move yet. Run YYCALIGNREC first.")
    (progn
      (yyc:msg (strcat "Which recorded move should profile \"" (yyc:profile) "\" use?"))
      (foreach x ps (setq i (1+ i)) (yyc:msg (strcat "    " (itoa i) ". " x "  (" (yyc:get-in x "AlignDate" "?") ")")))
      (setq ans (atoi (getstring "\nNumber: ")))
      (if (and (> ans 0) (<= ans (length ps)))
        (progn (yyc:set "AlignFrom" (nth (1- ans) ps))
               (yyc:msg (strcat "\"" (yyc:profile) "\" now uses the move from \"" (nth (1- ans) ps) "\".")))
        (yyc:msg "Not a choice."))
    )
  )
  (princ)
)

;;; ---------------------------------------------------------------------------
;;; Profiles window - YYCPROFILES (and YYCPROFILES): see, switch, add, edit, remove
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

(defun yyc:prof-show (idx / p users)
  (setq p (nth (atoi idx) *yyc-plist*))
  (setq users (vl-remove-if-not '(lambda (x) (and (/= x p) (= (strcase (yyc:align-src x)) (strcase p)))) *yyc-plist*))
  (set_tile "p0" (if (= (strcase p) (strcase (yyc:profile))) (strcat p "   (current profile)") p))
  (set_tile "p1" (strcat "Description:  " (yyc:get-in p "Description" "- none -")))
  (set_tile "p2" (strcat "Drawing no.:  " (yyc:get-in p "DwgNo" "-") "      Plot style:  " (yyc:get-in p "CTB" "-")))
  (set_tile "p3" (strcat "Grid:  " (yyc:fname "Grid" p)))
  (set_tile "p4" (strcat "Template:  " (yyc:fname "Template" p) "      Page setup:  " (yyc:get-in p "PageSetup" "-")))
  (set_tile "p5" (strcat "Layer reference:  " (yyc:fname "LayerRef" p) "      Layer map:  " (yyc:fname "LayerMap" p)))
  (set_tile "p6" (strcat "Kit folder:  " (yyc:short (yyc:get-in p "Kit" nil) 70)))
  (set_tile "p7" (strcat "Move:  " (yyc:move-text p)))
  (set_tile "p8" (if users (strcat "Its move is also used by:  " (yyc:join users ", ")) ""))
)

(defun c:YYCPROFILES ( / dcl f id res pick p name i)
  (setq *yyc-plist* (yyc:profiles))
  (setq dcl (vl-filename-mktemp "yycprof.dcl") f (open dcl "w"))
  (foreach x (list
    "yycprof : dialog { label = \"YYC profiles - one per project or area\";"
    "  : row {"
    "    : list_box { key = \"list\"; label = \"Profiles\"; width = 30; height = 15; }"
    "    : boxed_column { label = \"Details\";"
    "      : text { key = \"p0\"; width = 82; } : text { key = \"p1\"; width = 82; } : text { key = \"p2\"; width = 82; }"
    "      : text { key = \"p3\"; width = 82; } : text { key = \"p4\"; width = 82; } : text { key = \"p5\"; width = 82; }"
    "      : text { key = \"p6\"; width = 82; } : text { key = \"p7\"; width = 82; } : text { key = \"p8\"; width = 82; } } }"
    "  : text { value = \"Use = switch to it.  New = add one.  Edit = change its files (YYCSETUP).  Remove = take it off the list.\"; }"
    "  : row { : button { key = \"use\"; label = \"Use this profile\"; is_default = true; }"
    "          : button { key = \"new\"; label = \"New...\"; }"
    "          : button { key = \"edit\"; label = \"Edit...\"; }"
    "          : button { key = \"del\"; label = \"Remove\"; }"
    "          : button { key = \"cancel\"; label = \"Close\"; is_cancel = true; } }"
    "}") (write-line x f))
  (close f)
  (setq id (load_dialog dcl))
  (if (and (>= id 0) (new_dialog "yycprof" id))
    (progn
      (start_list "list")
      (foreach x *yyc-plist* (add_list (if (= (strcase x) (strcase (yyc:profile))) (strcat "* " x) (strcat "  " x))))
      (end_list)
      (setq i 0 pick "0")
      (foreach x *yyc-plist* (if (= (strcase x) (strcase (yyc:profile))) (setq pick (itoa i))) (setq i (1+ i)))
      (set_tile "list" pick)
      (yyc:prof-show pick)
      (action_tile "list" "(setq pick $value) (yyc:prof-show $value) (if (= $reason 4) (done_dialog 2))")
      (action_tile "use"  "(setq pick (get_tile \"list\")) (done_dialog 2)")
      (action_tile "new"  "(done_dialog 3)")
      (action_tile "edit" "(setq pick (get_tile \"list\")) (done_dialog 4)")
      (action_tile "del"  "(setq pick (get_tile \"list\")) (done_dialog 5)")
      (setq res (start_dialog))
      (unload_dialog id)
      (vl-file-delete dcl)
      (setq p (if (and pick (/= pick "")) (nth (atoi pick) *yyc-plist*)))
      (cond
        ((and (= res 2) p) (yyc:switch-profile p nil) (yyc:msg (strcat "Now using profile \"" p "\".")))
        ((= res 3)
         (setq name (yyc:ask "New profile name (e.g. ITB, 24C024 Area 4)" ""))
         (if (and name (/= name ""))
           (progn
             (yyc:switch-profile name (yyc:yes (strcat "Copy settings from \"" (yyc:profile) "\" to start with?") "Yes"))
             (yyc:msg (strcat "Created \"" name "\". Now set its own files:"))
             (c:YYCSETUP))))
        ((and (= res 4) p) (yyc:switch-profile p nil) (c:YYCSETUP))
        ((and (= res 5) p)
         (if (= (strcase p) (strcase (yyc:profile)))
           (yyc:msg "That's the current profile - switch to another one first, then remove it.")
           (if (yyc:yes (strcat "Remove \"" p "\" from the list?") "No")
             (progn (setenv "YYC_Profiles" (yyc:join (vl-remove p *yyc-plist*) "|"))
                    (yyc:msg (strcat "Removed \"" p "\"."))))))
      )
    )
    (yyc:msg "Could not open the profiles window.")
  )
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

(defun yyc:t-vplayers (doc / layers n name lo vp)
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
)
(defun c:YYCVPLAYERS () (yyc:run 'yyc:t-vplayers))

(defun yyc:t-vplock (doc / n vp)
  (setq n 0)
  (foreach pr (yyc:vp-list doc)
    (setq vp (cdr pr))
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

;; YYC Layer Reference as ((NAME colour LINETYPE lineweight) ...), read via ObjectDBX
(defun yyc:std-layers ( / path dbx res)
  (setq path (yyc:get "LayerRef" nil))
  (cond
    ((and *yyc-std* (= *yyc-std-path* path)) *yyc-std*)
    ((not (and path (findfile path))) nil)
    ((setq dbx (yyc:dbx-open path))
     (vlax-for L (vla-get-Layers dbx)
       (setq res (cons (list (strcase (vla-get-Name L)) (vla-get-Color L) (strcase (vla-get-Linetype L)) (vla-get-Lineweight L)) res)))
     (yyc:dbx-close dbx)
     (setq *yyc-std* res *yyc-std-path* path)
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

(defun yyc:layer-diff (L std / u pos ref out)
  (setq u (strcase (vla-get-Name L)) ref (assoc u std))
  (if (and (not ref) (setq pos (vl-string-position 45 u nil T))) (setq ref (assoc (substr u 1 pos) std)))
  (if ref
    (progn
      (if (/= (vla-get-Color L) (nth 1 ref))
        (setq out (cons (strcat "colour " (itoa (vla-get-Color L)) " should be " (itoa (nth 1 ref))) out)))
      (if (/= (strcase (vla-get-Linetype L)) (nth 2 ref))
        (setq out (cons (strcat "linetype " (vla-get-Linetype L) " should be " (nth 2 ref)) out)))
      (if (/= (vla-get-Lineweight L) (nth 3 ref))
        (setq out (cons (strcat "lineweight " (yyc:lw-text (vla-get-Lineweight L)) " should be " (yyc:lw-text (nth 3 ref))) out)))
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

;; --- CSV layer map ---
(defun yyc:read-map (path / f line parts res)
  (if (setq f (open path "r"))
    (progn
      (while (setq line (read-line f))
        (setq parts (yyc:split line ","))
        (if (and (= (length parts) 2) (/= (yyc:trim (car parts)) "") (/= (yyc:trim (cadr parts)) ""))
          (setq res (cons (cons (yyc:trim (car parts)) (yyc:trim (cadr parts))) res)))
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

(defun yyc:t-laymap (doc / layers csv map oldL newL moved del kept created)
  (setq layers (vla-get-Layers doc))
  (if (setq csv (yyc:need-file "LayerMap" "Select the layer map CSV (old,new per line)" "csv"))
    (progn
      (setq moved 0 del 0)
      (foreach p (yyc:read-map csv)
        (if (and (setq oldL (yyc:layer-ci layers (car p)))
                 (/= (strcase (vla-get-Name oldL)) (strcase (cdr p)))
                 (not (assoc (strcase (vla-get-Name oldL)) map)))
          (progn
            (if (not (setq newL (yyc:layer-ci layers (cdr p))))
              (setq newL (vla-Add layers (cdr p)) created (cons (cdr p) created)))
            (vla-put-Lock oldL :vlax-false)
            (setq map (cons (cons (strcase (vla-get-Name oldL)) (vla-get-Name newL)) map))
          )
        )
      )
      (if (assoc (strcase (vla-get-Name (vla-get-ActiveLayer doc))) map)
        (vla-put-ActiveLayer doc (vla-Item layers "0")))
      (vlax-for blk (vla-get-Blocks doc)
        (if (yyc:own-block-p blk) (vlax-for o blk (setq moved (+ moved (yyc:retag o map)))))
      )
      (foreach m map
        (if (yyc:err-p (vl-catch-all-apply '(lambda () (vla-Delete (vla-Item layers (car m))))))
          (setq kept (cons (car m) kept))
          (setq del (1+ del)))
        (yyc:msg (strcat "  " (car m) " -> " (cdr m)))
      )
      (yyc:msg (strcat "Moved " (itoa moved) " object(s); merged and removed " (itoa del) " layer(s)."))
      (if kept (yyc:msg (strcat "Emptied but not deleted (purge later): " (yyc:join kept ", "))))
      (if created (yyc:msg (strcat "CREATED with default properties - set them in LAYTRANS: " (yyc:join created ", "))))
    )
  )
)
(defun c:YYCLAYMAP () (yyc:run 'yyc:t-laymap))

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
  (foreach x details (yyc:r (strcat "         " x)))
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
                   bad main act ctb ctbbad xr h f rpt ff)
  (setq *yyc-rep* '() *yyc-flags* '())
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
            (if (= (yyc:layer-status cl std) "NON-STANDARD") (setq nonstd (cons cl nonstd)))
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
      (vlax-for lt (vla-get-Linetypes doc)
        (setq cl (strcase (vla-get-Name lt)))
        (if (and (not (member cl '("BYLAYER" "BYBLOCK" "CONTINUOUS"))) (not (vl-string-search "|" cl)) (not (member cl lins)))
          (setq ltbad (cons (vla-get-Name lt) ltbad))))
      (yyc:check "Linetypes come from CAA.LIN" (not ltbad) (yyc:list-line "Not in CAA.LIN" (reverse ltbad)))
    )
    (yyc:check "Linetypes come from CAA.LIN" T (list "Skipped - set the CAA .lin file in YYCSETUP"))
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
    (yyc:check "Layers match the YYC Layer Reference" (not (or nonstd diffs))
      (append (yyc:list-line "Not in the standard (modifiers allowed)" (reverse nonstd))
              (yyc:first-n (reverse diffs) 25)
              (if (or nonstd diffs) (list "Fix: LAYTRANS pass or YYCLAYMAP; full list via YYCLAYEXPORT"))))
    (yyc:check "Layers match the YYC Layer Reference" T (list "Skipped - set the layer reference in YYCSETUP"))
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
  (yyc:csv-replace (strcat (yyc:dfolder doc) "YYC_QA_Summary.csv") (yyc:dname doc) "Drawing,Checked,Flags,Flagged checks"
    (list (strcat (yyc:q (yyc:dname doc)) "," (yyc:stamp) "," (itoa (length *yyc-flags*)) "," (yyc:q (yyc:join (reverse *yyc-flags*) " | ")))))
  (yyc:msg (strcat "QA: " (if *yyc-flags* (strcat (itoa (length *yyc-flags*)) " flag(s)") "all automatic checks passed") " - summary in YYC_QA_Summary.csv"))
)
(defun c:YYCQA () (yyc:run 'yyc:t-qa))

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

(defun c:YYCFILELIST ( / folder files dwgno f bad)
  (setq folder (yyc:browse-folder "Delivery folder to list"))
  (if folder
    (progn
      (setq dwgno (strcase (yyc:ask "YYC drawing number" (yyc:get "DwgNo" "24C024"))))
      (setq files (yyc:dwgs-in folder))
      (if (setq f (open (strcat folder "\\filelist.txt") "w"))
        (progn (foreach x files (write-line x f)) (close f)))
      (foreach x files (if (not (yyc:name-ok (vl-filename-base x) dwgno)) (setq bad (cons x bad))))
      (yyc:msg (strcat (itoa (length files)) " drawing(s) written to " folder "\\filelist.txt"))
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


;;; ---------------------------------------------------------------------------
;;; Help window (YYCHELP), YYC menu, and ribbon loader
;;; ---------------------------------------------------------------------------

;; (step  command  what-it-does). AutoCAD's own commands are marked (AutoCAD).
(setq *yyc-help*
 '(("Setup" "YYCSETUP"      "Create or edit a profile: one set of kit files (grid, template, layer reference) per project or area. Do this first.")
   ("Setup" "YYCPROFILES"   "Profiles window: see every profile with its description, files and recorded move; switch, add, edit or remove one.")
   ("2 Name" "YYCRENAME"    "Rename every DWG in a folder to drawing no. + sheet no. Shows a preview first. Do it before opening the files.")
   ("3 Sheet" "YYCPAGESETUP" "Apply the YYC - Titleblock page setup from the template to every layout, giving a true A0.")
   ("3 Sheet" "YYCTITLEBLOCK" "Experimental: swap the exported titleblock for the YYC A0 titleblock and carry the title text into its fields.")
   ("4 Sheet" "YYCSCHEDULES" "Click a schedule viewport: everything it shows moves to paper space at the same size and spot, then it offers to delete the empty viewport. One at a time. Do this before moving the plan.")
   ("4 Sheet" "CHSPACE"     "(AutoCAD) For anything YYCSCHEDULES can't take, e.g. objects crossing a viewport edge.")
   ("5 Grid" "YYCGRIDIN"    "Insert the YYC grid for this area at 0,0,0 as one block. You pick which grid.")
   ("5 Grid" "YYCALIGNREC"  "First sheet: 4 clicks like ALIGN (source 1, destination 1, source 2, destination 2). Moves model space, never scales, saves the move to the profile.")
   ("5 Grid" "YYCALIGNAPPLY" "Every other sheet: apply the saved move with no clicks, and turn the viewports to match. Will not move a drawing twice.")
   ("7 Viewport" "YYCVPFOLLOW" "Turn the drawing viewport so the sheet shows the same plan as the export, project north up. The align tools do this for you; use it only if you skipped it.")
   ("5 Grid" "YYCALIGNUSE"  "Let this profile use a move recorded in another profile (same export location).")
   ("6 Clean" "YYCGRIDOUT"  "Remove the grid block once the plan is on the grid.")
   ("6 Clean" "YYCCLEAN"    "Delete empty text, purge everything 3 times, audit and fix.")
   ("6 Clean" "YYCFIND0"    "List names containing $0$ left over from binding xrefs.")
   ("7 Viewport" "YYCVPLAYERS" "Put each viewport on its own VIEWPORT# no-plot layer.")
   ("7 Viewport" "YYCVPLOCK" "Lock every viewport and list its scale. Only after the scale and rotation are right.")
   ("8 Layers" "YYCLWDEFAULT" "Set every layer's lineweight to Default (fixes the Revit LineWeight009/025/030).")
   ("8 Layers" "LAYTRANS"   "(AutoCAD) Layer Translator - one pass per job, Map Same first, force BYLAYER in Settings.")
   ("8 Layers" "YYCLAYMAP"  "Merge layers from a saved CSV map (old,new). For mappings you repeat on every sheet.")
   ("8 Layers" "YYCLAYEXPORT" "Write all layers to YYC_LayerExport.csv, flagged against the YYC Layer Reference.")
   ("8 Layers" "SETBYLAYER" "(AutoCAD) Force colour and linetype back to BYLAYER. Ctrl+A first, include blocks.")
   ("8 Layers" "YYCZERO"    "Show and select what is on 0 / DEFPOINTS. Never moves anything.")
   ("8 Layers" "LAYWALK"    "(AutoCAD) Walk every layer to check its contents. Tick Restore on exit.")
   ("9 QA" "YYCFINAL"       "Lock viewports, make the main layout active, zoom extents.")
   ("9 QA" "YYCQA"          "CADD Manual 5.15 check. Reports only - writes <drawing>_YYC-QA.txt and YYC_QA_Summary.csv.")
   ("10 Send" "YYCFILELIST" "Write filelist.txt for the File Description and flag badly named files.")
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

;; "YYC" drop-down menu on the menu bar, built through ActiveX each session
(defun yyc:build-menu ( / acad mg menus pm sub step)
  (setq acad (yyc:acad) mg (yyc:try 'vla-Item (list (vla-get-MenuGroups acad) 0)))
  (if (and mg (not (yyc:try 'vla-Item (list (vla-get-Menus mg) "YYC"))))
    (progn
      (setq menus (vla-get-Menus mg) pm (vla-Add menus "YYC"))
      (vla-AddMenuItem pm 0 "Help - all commands in order..." "^C^CYYCHELP ")
      (vla-AddSeparator pm 1)
      (foreach h *yyc-help*
        (if (/= (car h) step)
          (setq step (car h) sub (vla-AddSubMenu pm (vla-get-Count pm) step)))
        (vla-AddMenuItem sub (vla-get-Count sub) (cadr h) (strcat "^C^C" (cadr h) " "))
      )
      (vl-catch-all-apply 'vla-InsertInMenuBar (list pm (vla-get-Count (vla-get-MenuBar acad))))
    )
  )
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
  '(("Titleblock" . yyc:t-titleblock) ("Align" . yyc:t-alignapply) ("Pagesetup" . yyc:t-pagesetup) ("Vplayers" . yyc:t-vplayers) ("Lwdefault" . yyc:t-lwdefault)
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
(vl-catch-all-apply 'yyc:build-menu nil)
(vl-catch-all-apply 'yyc:load-ribbon nil)
(princ (strcat "\nYYC CAD Tools v" *yyc-version* " (ActiveX) loaded. Type YYCHELP for every command in order. Profile: " (yyc:profile) ". Commands: YYCHELP YYCPROFILES YYCRENAME YYCTITLEBLOCK YYCPAGESETUP YYCSCHEDULES YYCGRIDIN YYCALIGNREC YYCALIGNAPPLY YYCVPFOLLOW YYCALIGNUSE YYCGRIDOUT YYCCLEAN YYCFIND0 YYCVPLAYERS YYCVPLOCK YYCLWDEFAULT YYCLAYMAP YYCLAYEXPORT YYCZERO YYCQA YYCFINAL YYCFILELIST YYCBATCH"))
(princ)
