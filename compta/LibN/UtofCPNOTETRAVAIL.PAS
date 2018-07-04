{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/11/2006
Modifié le ... : 13/06/2007
Description .. : Source TOF de la FICHE : CPNOTETRAVAIL ()
Mots clefs ... : TOF;CPNOTETRAVAIL
*****************************************************************}
Unit uTOFCPNOTETRAVAIL ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,       //Agllancefiche
     EdtREtat,       // LanceEtat
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     UtileAgl,    //LanceEtat
     eMul,
     MainEAgl,      //AglLancefiche
{$ENDIF}
     forms,
     utob,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     hpanel,          // THPanel
     Grids,
     ent1,           //VH^
     CPREVDocTravaux_TOF,   // CPLanceFiche_CPRevDocTravaux
     Windows,        //Formkeydown
     HTB97,          //TToolBarButton97
     Menus,          //TPopupMenu
     HrichOLE,       //THRichEditOle
     Dialogs,        //TSaveDialog
     Hxlspas,        //ExportGrid
     Formule,        //Gformule
     Clipbrd,        //clipboard
     UTOF ;

const
//les colonnes du tableau
  clFixe                : integer = 0;
  clDate                : integer = 1;
  clLib                 : integer = 2;
  clMt1                 : integer = 4;
  clMt2                 : integer = 5;
  clSens                : integer = 3;      //FQ19958
  clOpe                 : integer = 6;
  clDeb                 : integer = 7;
  clCre                 : integer = 8;
  clBN                  : integer = 9;

  StrErrorClpbrd        : Array [0..6] of String = (
                          'Import réussi.',
                          'Le contenu du presse-papier ne convient pas.',
                          'Le presse-papier est vide.',
                          'Le code exercice n''est pas identique.',
                          'Le compte général n''est pas identique.',
                          'L''écriture n''est pas au bon format.',
                          'Le descriptif de l''écriture ne convient pas.'
                          );

Type
  TOF_CPNOTETRAVAIL = Class (TOF)
    private

    CpGeneral           : String;               // compte général
    CodeExercice        : String;               // Code exercice
    StrLibCpt           : String;               // Libellé compte général
    DbTotCE             : Double;               // total solde crédit Exercice
    DbTotDE             : Double;               // total solde débit Exercice
    DbTotCP             : Double;               // total solde débit n-1
    DbTotDP             : Double;               // total solde crédit n-1
    DateFinExercice     : String;               // date de fin d'exercice
    DateDebExercice     : String;               // date de début d'exercice
    CpSens              : String;               // Sens comptable
    HorsFListe          : Boolean;              // saisie dans la liste ou non
    NotePrec            : Boolean;              // visu de la note exercice précédent
    TriDateDesc         : Boolean;              // est un tri descendant sur date
    TriLibDesc          : Boolean;              // est un tri descendant sur libellé
    TriMtDesc           : Boolean;              // est un tri descendant sur montant
    CpCycleRevision     : String;               // Cycle de révision

    procedure RechercheDonnees;
    function ValideLigne (LaLigne : integer ; SansCalc : Boolean = True ; SurSortie : Boolean = False) : Boolean;
    procedure AjouteLigne (AuDessus : Boolean = False ; Premiere : Boolean = False ; LG0 : Boolean = False);
    procedure BinsertOnClick (Sender : TObject);
    procedure FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure SortieSansSauvegarde;
    procedure SortieSansSauvegardeOnClick (Sender : TObject);
    procedure EnregTableau (Sortie : Boolean = True);
    procedure BValideOnClick (Sender : TObject);
    procedure FListeCellEnter (Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure FListeCellExit (Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure CalculTotaux;
    procedure SupprimeLigne;
    procedure BSupprimeLigneOnClick (Sender : TObject);
    procedure BDeleteOnCLick (Sender : Tobject);
    procedure Repro (RTotal : Boolean = True);
    procedure CommMillesime;
    procedure CommMillesimeOnClick (Sender : Tobject);
    procedure CommCompte;
    procedure CommcompteOnClick (Sender : Tobject);
    procedure BBlocNoteOnCLick (Sender : Tobject);
    procedure HPBOnClose (Sender : Tobject);
    procedure FListeOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure FlisteOnRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure SaisieBNOnExit (Sender : TObject);
    procedure SaisieBNOnEnter (Sender : Tobject);
    procedure SaisieTitreOnExit (Sender : TObject);
    procedure SaisieTitreOnEnter (Sender : Tobject);
//    function GetFormulePourCalc(Formule : hstring) : Variant ;
    procedure BExportOnClick (Sender : TObject);
//    function CalculCellule (LaCol, LaLigne : integer) : Boolean;
    procedure OnClickCollerEcriture( Sender : TObject );
    function ImportClipBoard (EstCumul : Boolean = False) : integer;
    procedure ImprimeTableau;
    procedure BImprimeTableauOnCLick (Sender : TObject);
//    procedure VisuAncienneNote;         //FQ19960
    procedure BZoomOnClick (Sender : Tobject);
    function YaNoteAvant : boolean;
    // combochanged
//    procedure ComboChanged (Sender : Tobject);
//    procedure ComboKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure BlocNoteOnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);

    procedure ReproTotOnClick (Sender : Tobject);
    procedure ReporElonClick (Sender : Tobject);
    procedure RenomerNoteOnClick (Sender : Tobject);
    procedure ObsNoteOnClick (Sender : Tobject);
    procedure BlocnoteOnclickF11 (Sender : Tobject);
    procedure TridateOnClick (Sender : Tobject);
    procedure TriMontantOnClick (Sender : Tobject);
    procedure TriLibelleOnClick (Sender : tobject);
    procedure LeTri (LeType : String) ;
    procedure CycleCommentaireOnClick (Sender : Tobject);
    procedure CycleObjectifOnClick (Sender : Tobject);
    procedure CycleSyntheseOnClick (Sender : Tobject);
    procedure DeplaceBlocNote;
    procedure HPBOnVisibleChanged (Sender : Tobject);
    procedure PgNoteOnChanged (Sender : Tobject);
    procedure MajBlocNote (LaLigne : integer);

    public
    //Grille de saisie
    FListe              : THGrid;
    BValide             : TToolbarButton97;
    BInsert             : TToolbarButton97;
    BSupprime           : TToolbarButton97;
    BDelete             : TToolbarButton97;
    BFerme              : TToolbarButton97;
    BExport             : TToolbarButton97;
    BImprime            : TToolbarButton97;
    BZoom               : TToolbarButton97;
    BCommentaire        : TToolbarButton97;

    MenuCom             : TPopupMenu;
    PopF11              : TPopupMenu;
    PopTri              : TPopupMenu;
    PopNotePrec         : TPopupMenu;           //FQ19960
    TExercice           : THLabel;
    PGNote              : TPageControl;
//    ComboPlusMoins      : THValComboBox ;

    //les totaux
    TTotDeb             : THLabel;              // total débit
    TTotCre             : THLabel;              // total crédit
    TSoldeDeb           : THLabel;              // Solde comptable débit
    TSoldeCre           : THLabel;              // Solde comptable crédit
    TPDeb               : THLabel;              // Pourcentage débit
    TPCre               : THLabel;              // Pourcentage crédit
    TBalDeb             : ThLabel;              // Total Balance débit
    TBalCre             : ThLabel;              // Total Ballance Crédit

    //le bloc notes
    BBlocNote           : TToolBarButton97;
    FBlocNote           : THRichEditOle;
    HPB                 : TToolWindow97;

    //la page informations
    SaisieBN            : THRichEditOLE;
    SaisieTitre         : THEdit;

    ASaveDialog         : TSaveDialog;          // fenetre de dialogue pour export

    TobNoteTravail      : Tob;

    procedure OnArgument (S : string ) ; override ;
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

procedure CPLanceFiche_CPNOTETRAVAIL( vStCompte : string; vBoPrecedent : Boolean = False; CodeExePrec : String = '' );

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Vierge,  // TFFiche
  AGLInit, // ActionToString
  uLibRevision,
  TntStdCtrls;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 10/11/2006
Modifié le ... : 13/06/2007
Description .. : Appel de la fiche
Suite ........ : GCO - 13/06/2007 - Ajout du vStActionFiche
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_CPNOTETRAVAIL( vStCompte : string; vBoPrecedent : Boolean = False; CodeExePrec : String = '' );
begin
  AGLLanceFiche('CP', 'CPNOTETRAVAIL', '', '', vStCompte + ';' + BoolToStr(vBoPrecedent) + ';' + CodeExePrec);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 13/11/2006
Modifié le ... :   /  /
Description .. : initialisation
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.OnArgument( S : String );
var
  Q                     : TQuery;
  lBoPrecedent          : Boolean;
  tdeb                  : extended;
  tcre                  : extended;
  tsolde                : extended;
  i,j                   : integer;
  MenuI                 : TMenuItem;
  CodeExePrec           : String;
  AncienneRevision      : integer; // FQ 20579
  lStEtatCycle          : String;
  lMenu                 : TMenuItem;

begin
  Inherited ;

  //les controles
  PGNote                := TPageControl(GetControl('PG1'));
  FListe                := THGrid(GetControl('GRVAR'));
  TExercice             := THLabel(GetControl('TEXERCICE'));
  BValide               := TToolbarButton97(GetControl('BVALIDER'));
  BInsert               := TToolbarButton97(GetControl('BINSERT'));
  BSupprime             := TToolbarButton97(GetControl('BSUPPRIME'));
  BDelete               := TToolbarButton97(GetControl('BDELETE'));
  BFerme                := TToolbarButton97(GetControl('BFERME'));
  BBlocNote             := TToolbarButton97(GetControl('BBLOCNOTE'));
  BExport               := TToolbarButton97(GetControl('BEXPORT'));
  BImprime              := TToolbarButton97(GetControl('BIMPRIMER'));
  BZoom                 := TToolbarButton97(GetControl('BZOOM'));
  BCommentaire          := TToolbarButton97(GetControl('BCOMMENTAIRE'));
  MenuCom               := TPopupMenu(GetControl('PM_COMM'));
  PopF11                := TPopupMenu(GetControl('POPF11'));
  PopTri                := TPopupMenu(GetControl('POPTRI'));
  PopNotePrec           := TPopupMenu(GetControl('POPNOTEPREC'));     //FQ19960
  TTotDeb               := THLabel(GetControl('TTOTDEB'));
  TTotCre               := THLabel(GetControl('TTOTCRE'));
  TSoldeDeb             := THLabel(GetControl('TSOLDEDEB'));
  TSoldeCre             := THLabel(GetControl('TSOLDECRE'));
  TPDeb                 := THLabel(GetControl('TPDEB'));
  TPCre                 := THLabel(GetControl('TPCRE'));
  TBalDeb               := THLabel(GetControl('TBALDEB'));
  TBalCre               := THLabel(GetControl('TBALCRE'));
  SaisieBN              := THRichEditOLE(GetControl('SAISIEBN'));
  SaisieTitre           := THEdit(GetControl('SAISIETITRE'));
//  ComboPlusMoins        := THValComboBox(GetControl('PLUSMOINS', True));

  CPGeneral             := ReadTokenSt(S);
  lBoPrecedent          := StrToBool(ReadToKenST(S));
  CodeExePrec           := ReadToKenSt(S);
  NotePrec              := lBoPrecedent;

  TriDateDesc           := True;
  TriLibDesc            := True;
  TriMtDesc             := True;

  AncienneRevision      := 0;

  try
    Q                   := OpenSQL('SELECT G_LIBELLE, G_SENS, ' +
                          'G_TOTDEBE, G_TOTCREE, G_TOTDEBP, G_TOTCREP, G_CYCLEREVISION ' +
                          'FROM GENERAUX WHERE G_GENERAL = "' + CPGeneral + '"', True);

    if not Q.Eof then
    begin
      StrLibCpt         := Q.FindField('G_LIBELLE').AsString;
      CpSens            := Q.FindField('G_SENS').AsString;
      DbTotDE           := Q.FindField('G_TOTDEBE').AsFloat;
      DbTotCE           := Q.FindField('G_TOTCREE').AsFloat;
      DbTotDP           := Q.FindField('G_TOTDEBP').AsFloat;
      DbTotCP           := Q.FindField('G_TOTCREP').AsFloat;
      CpCycleRevision   := Q.FindField('G_CYCLEREVISION').AsString;
    end;

    if lBoPrecedent then
    begin
      CodeExercice      := CodeExePrec;
      // GCO - 26/06/2007 - pas d'accès à CPREVDOCTRAVAUX sur le précédent
      BCommentaire.Visible := False;
    end  
    else
    begin
      CodeExercice      := VH^.EnCours.Code;
    end;

    // GCO - 13/06/2007 - FQ 20480
    // GCO - 11/07/2007 - FQ 20946 - Blocage en consultation suivant l'état de la révision
    if CpCycleRevision <> '' then
    begin
      lStEtatCycle := GetColonneSQL('CREVCYCLE', 'CCY_ETATCYCLE',
                      'CCY_CODECYCLE = "' + CpCycleRevision + '" AND  ' +
                      'CCY_EXERCICE = "' + CodeExercice  + '"');

      if (VH^.Revision.DossierSupervise) or
         (VH^.Revision.DossierPretSupervise and (not JaileRoleCompta(rcSuperviseur))) or
         ((lStEtatCycle = cValide) and (not JaileRoleCompta(rcSuperviseur))) or
         (lStEtatCycle = cSupervise) then
        TFVierge(Ecran).TypeAction := TaConsult;
    end;
    // FIN GCO

  finally
    Ferme(Q);
  end;



  // Création de la fenêtre BLOC NOTE
  HPB                   := TToolWindow97.Create(nil);
  HPB.Parent            := Ecran;
  HPB.Caption           := 'Notes sur la ligne';
  HPB.Visible           := False;
  HPB.Left              := 350;
  HPB.Top               := 260;
  HPB.Height            := 140;
  HPB.width             := 250;
  { FQ 21242 BVE 29.08.07 }
  // On ne peut pas modifier la note précédente
  HPB.Enabled           := not(NotePrec);
  { END 21242 }

  FBlocNote             := THRichEditOle.Create(nil);
  FBlocNote.Parent      := HPB;
  FBlocNote.Align       := AlClient; 
  { FQ 20826 BVE 25.06.07 }
  FBlocNote.Font.Height := -11;
  { END FQ 20826 }
  FBlocNote.Clear;
  FBlocNote.OnKeyDown   := BlocNoteOnKeyDown;
  { FQ 21242 BVE 29.08.07 }
  // On ne peut pas modifier la note précédente
  FBlocNote.ReadOnly    := NotePrec;
  { END 21242 }

//  FblocNote.ReadOnly    := True;      //Consultation uniquement

  //Boite dialogue pour export
  ASaveDialog                 := TSaveDialog.Create(Ecran);
  ASaveDialog.Filter          := 'Fichier Texte (*.txt)|*.txt|Fichier Excel (*.xls)|*.xls|Fichier Ascii (*.asc)|*.asc|Fichier Lotus (*.wks)|*.wks|Fichier HTML (*.html)|*.html|Fichier XML (*.xml)|*.xml';
  ASaveDialog.DefaultExt      := 'XLS';
  ASaveDialog.FilterIndex     := 1;
  ASaveDialog.Options         := ASaveDialog.Options + [ofOverwritePrompt, ofPathMustExist, ofNoReadonlyReturn, ofNoLongNames] - [ofEnableSizing];

  // les proprietes
  Ecran.OnKeyDown             := FormKeyDown;
  BInsert.OnClick             := BinsertOnClick;
  BFerme.OnClick              := SortieSansSauvegardeOnClick;
  BValide.OnClick             := BValideOnClick;
  BSupprime.OnClick           := BSupprimeLigneOnClick;
  BDelete.Visible             := True;
  BDelete.OnClick             := BDeleteOnCLick;
  BImprime.Visible            := True;
  BImprime.OnClick            := BImprimeTableauOnCLick;
  BBlocNote.OnClick           := BBlocNoteOnCLick;
  BExport.OnClick             := BExportOnClick;
//  BZoom.OnClick               := BZoomOnClick;      //FQ19960
  BZoom.Visible               := False;
  HPB.OnClose                 := HPBOnClose;
  HPB.OnVisibleChanged        := HPBOnVisibleChanged;
  FListe.OnCellEnter          := FListeCellEnter;
  FListe.OnCellExit           := FListeCellExit;
  FListe.OnRowEnter           := FListeOnRowEnter;
  Fliste.OnRowExit            := FlisteOnRowExit;
  TForm(Ecran).ShowHint       := True;
  SaisieBN.OnExit             := SaisieBNOnExit;
  SaisieBN.OnEnter            := SaisieBNOnEnter;
  SaisieTitre.OnExit          := SaisieTitreOnExit;
  SaisieTitre.OnEnter         := SaisieTitreOnEnter;
  PGNote.OnChange             := PgNoteOnChanged;
//  ComboPlusMoins.OnChange     := ComboChanged;
//  ComboPlusMoins.OnKeyDown    := ComboKeyDown;

  // GCO - 02/08/2007 - FQ 21226
  SaisieTitre.MaxLength := 35;

  //menu pop commentaires
  if Assigned(MenuCom) then
  begin
    for i := 0 to MenuCom.Items.Count -1 do
    begin
      if MenuCom.Items[i].Name = 'COMCYCLE' then
        if CpCycleRevision = '' then
          MenuCom.Items[i].Visible  := False
        else
          MenuCom.Items[i].OnClick  := CycleCommentaireOnClick;
      if MenuCom.Items[i].Name = 'OBJREV' then
        if CpCycleRevision = '' then
          MenuCom.Items[i].Visible  := False
        else
          MenuCom.Items[i].OnClick  := CycleObjectifOnClick;
      if MenuCom.Items[i].Name = 'SYNTHCYCLE' then
        if CpCycleRevision = '' then
          MenuCom.Items[i].Visible  := False
        else
          MenuCom.Items[i].OnClick  := CycleSyntheseOnClick;
      if MenuCom.Items[i].Name = 'COM' then
        MenuCom.Items[i].OnClick  := CommMillesimeOnClick;
      if MenuCom.Items[i].Name = 'COC' then
        MenuCom.Items[i].OnClick  := CommcompteOnClick;
    end;
  end;

  //menu pop Tri
  if Assigned(PopTri) then
  begin
    for i := 0 to PopTri.Items.Count -1 do
    begin
      if PopTri.Items[i].Name = 'TRIDATE' then
        PopTri.Items[i].OnClick := TriDateOnClick;
      if PopTri.Items[i].Name = 'TRILIBELLE' then
        PopTri.Items[i].OnClick := TriLibelleOnClick;
      if PopTri.Items[i].Name = 'TRIMONTANT' then
        PopTri.Items[i].OnClick := TriMontantOnClick;
    end;
  end;

  //menu pop contextuel
  if Assigned(PopF11) then
  begin
    for i := 0 to PopF11.Items.Count -1 do
    begin
      if PopF11.Items[i].Name = 'ALG' then
      begin
        PopF11.Items[i].OnClick := BinsertOnClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;

      if PopF11.Items[i].Name = 'SLG' then
      begin
        PopF11.Items[i].OnClick := BSupprimeLigneOnClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;  

      if PopF11.Items[i].Name = 'ANC' then
        if (YaNoteAvant) and (not NotePrec) then
        begin
          { FQ 20579 BVE 12.06.07
          PopF11.Items[i].OnClick := BZoomOnClick; }
          AncienneRevision := i;
          { END FQ 20579 }
        end
        else
        begin
          PopF11.Items[i].Visible := False;
          PopF11.Items[i].OnClick := nil;
        end;

      if PopF11.Items[i].Name = 'IMP' then
        PopF11.Items[i].OnClick := BImprimeTableauOnCLick;

      if PopF11.Items[i].Name = 'VAL' then
        PopF11.Items[i].OnClick := BValideOnClick;

      if PopF11.Items[i].Name = 'EXT' then
        PopF11.Items[i].OnClick := SortieSansSauvegardeOnClick;

      if PopF11.Items[i].Name = 'DUP' then
      begin
        PopF11.Items[i].OnClick := ReproTotOnClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;

      if PopF11.Items[i].Name = 'COP' then
      begin
        PopF11.Items[i].OnClick := ReporElonClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;

      if PopF11.Items[i].Name = 'REN' then
      begin
        PopF11.Items[i].OnClick := RenomerNoteOnClick;
        PopF11.Items[i].Visible := TFVierge(Ecran).TypeAction <> TaConsult;
      end;  

      if PopF11.Items[i].Name = 'OBS' then
        PopF11.Items[i].OnClick := ObsNoteOnClick;

      if PopF11.Items[i].Name = 'NOT' then
        PopF11.Items[i].OnClick := BlocnoteOnclickF11;


      if PopF11.items[i].Name = 'DOCTRAVAUXF11' then
      begin
        // GCO - 24/09/2007 - FQ 21242
        PopF11.items[i].Visible := BCommentaire.Visible;

        for j := 0 to PopF11.Items[i].Count -1 do
        begin
          lMenu := PopF11.Items[i].Items[j];

          // GCO - 26/07/2007 - FQ 20976
          if lMenu.Name = 'COMCYCLEF11' then
          begin
            lMenu.Visible := CpCycleRevision <> '';
            lMenu.OnClick := CycleCommentaireOnClick;
          end;

          if lMenu.Name = 'OBJREVF11' then
          begin
            lMenu.Visible := CpCycleRevision <> '';
            lMenu.OnClick := CycleObjectifOnClick;
          end;

          if lMenu.Name = 'SYNTHCYCLEF11' then
          begin
            lMenu.Visible := CpCycleRevision <> '';
            lMenu.OnClick := CycleSyntheseOnClick;
          end;
          // FIN GCO

          if lMenu.Name = 'COMM' then
          begin
            // GCO - 26/06/2007 - FQ 20845
            lMenu.Visible := BCommentaire.Visible;
            lMenu.OnClick := CommMillesimeOnClick;
          end;

          if lMenu.Name = 'COMC' then
          begin
            // GCO - 26/06/2007 - FQ 20845
            lMenu.Visible := BCommentaire.Visible;
            lMenu.OnClick := CommcompteOnClick;
          end;
        end;  
      end;

      // GCO - 04/09/2007 - FQ 21068
      if PopF11.Items[i].Name = 'COLLERECR' then
      begin
        PopF11.Items[i].Visible := Clipboard.HasFormat(CF_TEXT) and
                                   (TFVierge(Ecran).TypeAction <> TaConsult);
        PopF11.Items[i].OnClick := OnClickCollerEcriture;
      end;  
    end;
  end;

  ActivateXpPopUp( PopF11 );

  //menu pop NotePrec   //FQ19960
  if (Assigned(PopNotePrec)) and (not lBoPrecedent) then
  begin
    Try
      Q                 := OpenSQL('SELECT EX_EXERCICE, EX_LIBELLE, EX_DATEFIN, EX_DATEDEBUT FROM EXERCICE ' +
                                   'WHERE EX_EXERCICE IN (SELECT CNO_EXERCICE FROM CPNOTETRAVAIL ' +
                                   ' WHERE CNO_GENERAL="' + CpGeneral + '")', True);
      While not Q.Eof do
      begin
        if Q.FindField('EX_EXERCICE').AsString <> CodeExercice then
        begin
          MenuI         := TMenuItem.Create(PopNotePrec);
          MenuI.Name    := 'EX' + Q.FindField('EX_EXERCICE').AsString;
          MenuI.Caption := Q.FindField('EX_LIBELLE').AsString ;
          MenuI.Hint    := 'Du ' + Q.FindField('EX_DATEDEBUT').AsString + ' au ' + Q.FindField('EX_DATEFIN').AsString;
          MenuI.Tag     := 0;
          MenuI.OnClick := BZoomOnClick;
          MenuI.OnDrawItem  := PopTri.Items[0].OnDrawItem;
          MenuI.OnMeasureItem := PopTri.Items[0].OnMeasureItem;
          PopNotePrec.Items.Add(MenuI);
          { FQ 20579 BVE 12.06.07 }
          if AncienneRevision > 0 then
          begin
             MenuI         := TMenuItem.Create(PopNotePrec);
             MenuI.Name    := 'E2' + Q.FindField('EX_EXERCICE').AsString;
             MenuI.Caption := Q.FindField('EX_LIBELLE').AsString ;
             MenuI.Hint    := 'Du ' + Q.FindField('EX_DATEDEBUT').AsString + ' au ' + Q.FindField('EX_DATEFIN').AsString;
             MenuI.Tag     := 0;
             MenuI.OnClick := BZoomOnClick;
             MenuI.OnDrawItem  := PopTri.Items[0].OnDrawItem;
             MenuI.OnMeasureItem := PopTri.Items[0].OnMeasureItem;
             PopF11.Items[AncienneRevision].Add(MenuI);
          end
          { END FQ 20579 }
        end;
        Q.Next;
      end;

      if PopNotePrec.Items.Count > 0 then
        bzoom.Visible   := True;
    Finally
      Ferme(Q);
    end;
  end;

  // l'entete
  TForm(Ecran).Caption  := 'Note de travail : ' + CpGeneral + ' - ' + StrLibCpt;
  UpdateCaption(TForm(Ecran));
  Q                     := OpenSQL('SELECT EX_LIBELLE, EX_DATEFIN, EX_DATEDEBUT FROM EXERCICE WHERE EX_EXERCICE="' + CodeExercice + '"', True);
  DateFinExercice       := Q.Fields[1].AsString;
  DateDebExercice       := Q.Fields[2].AsString;
  TExercice.Caption     := 'Solde de fin d''exercice : ';
  Ferme(Q);


    // les parametres
  FListe.ColCount               := 10;
  FListe.RowCount               := 2;
  FListe.FixedCols              := 1;
  FListe.FixedRows              := 1;
  FListe.ColWidths[clFixe]      := 10;
  FListe.ColWidths[clDate]      := 70;
  FListe.ColWidths[clLib]       := 120;
  FListe.ColWidths[clMt1]       := 70;
  FListe.ColWidths[clMt2]       := 70;
  FListe.ColWidths[clSens]      := 65;
  FListe.ColWidths[clOpe]       := 65;
  FListe.ColWidths[clDeb]       := 90;
  FListe.ColWidths[clCre]       := 90;
  FListe.ColWidths[clBN]        := 40;

  FListe.TwoColors              := True;
  FListe.ColTypes[clDate]       := 'D';
  FListe.ColFormats[clDate]     := 'dd/mm/yyyy';  //sinon effet de bord // ??
//  FListe.ColFormats[clDate]     := ShortDateFormat;

  FListe.ColTypes[clLib]        := 'S';
  { FQ 20645 BVE 12.06.07 }
  FListe.ColLengths[clLib]      := 35;
  { END FQ 20645 }

  FListe.ColTypes[clMt1]        := 'K';
//  FListe.ColFormats[clMt1]      := '##,##0.00';
  FListe.ColFormats[clMt1]      := 'R';
  FListe.ColAligns[clMt1]       := taRightJustify;

  FListe.ColTypes[clMt2]        := 'K';
//  FListe.ColFormats[clMt2]      := '##,##0.00';
  FListe.ColAligns[clMt2]       := taRightJustify;

  FListe.CalcInCell             := True;          // pour permettre le calcul dans les cellules

  FListe.ColTypes[clDeb]        := 'R';
  FListe.ColFormats[clDeb]      := '##,##0.00';
  FListe.ColAligns[clDeb]       := taRightJustify;

  FListe.ColTypes[clCre]        := 'R';
  FListe.ColFormats[clCre]      := '##,##0.00';
  FListe.ColAligns[clCre]       := taRightJustify;

  FListe.ColFormats[clSens]     := 'CB=TTSENS|AND CO_CODE <> "M"|';
  FListe.ColAligns[clSens]      := taCenter;

  FListe.ColFormats[clOpe]      := 'CB=CPOPERATEUR||';
  FListe.ColAligns[clOpe]       := taCenter;
  FListe.ColEditables[clDate]   := True;
  FListe.ColEditables[clLib]    := True;
  FListe.ColEditables[clMt1]    := True;
  FListe.ColEditables[clMt2]    := True;
  FListe.ColEditables[clSens]   := True;
  FListe.ColEditables[clOpe]    := True;
  FListe.ColEditables[clDeb]    := False;
  FListe.ColEditables[clCre]    := False;
  FListe.ColAligns[clBN]        := taCenter;
  FListe.ColEditables[clBN]     := False;
  FListe.Options                := FListe.Options - [goRowSelect] + [goEditing] + [goTabs];



  FListe.Cells[clDate, 0]       :=  TraduireMemoire('Date');
  FListe.Cells[clLib, 0]        := TraduireMemoire('Libellé');
  FListe.Cells[clMt1, 0]        := TraduireMemoire('Montant');
  FListe.Cells[clMt2, 0]        := TraduireMemoire('Montant 2');
  FListe.Cells[clSens, 0]       := TraduireMemoire('Sens');
  FListe.Cells[clOpe, 0]        := TraduireMemoire('Opérateur');
  FListe.Cells[clDeb, 0]        := TraduireMemoire('Débit');
  FListe.Cells[clCre, 0]        := TraduireMemoire('Crédit');
  FListe.Cells[clBN, 0]         := TraduireMemoire('Notes');

  PGNote.ActivePage             := PGNote.Pages[0];

    TobNoteTravail := Tob.Create('CPNOTETRAVAIL', nil, -1);

  //initialisation
  tdeb                  := DbTotDE;
  tcre                  := DbTotCE;
  if lBoPrecedent then
  begin
    tdeb                := DbTotDP;
    tcre                := DbTotCP;
  end;

//  TSoldeDeb.Caption := STRFMONTANT(tdeb, 15, V_PGI.OkDecV, '', True);
//  TSoldeCre.Caption := STRFMONTANT(tcre, 15, V_PGI.OkDecV, '', True);
  if tdeb >= tcre then
  begin
    tsolde              := tdeb - tcre;
    TSoldeDeb.Caption   := STRFMONTANT(tsolde, 15, V_PGI.OkDecV, '', True);
    TSoldeCre.Caption   := '';
    TExercice.Caption   := TExercice.Caption + STRFMONTANT(tsolde, 15, V_PGI.OkDecV, '', True) +  ' D au ' + DateFinExercice
  end
  else
  begin
    tsolde              := tcre - tdeb;
    TSoldeCre.Caption   := STRFMONTANT(tsolde, 15, V_PGI.OkDecV, '', True);
    TSoldeDeb.Caption   := '';
    TExercice.Caption   := TExercice.Caption + STRFMONTANT(tsolde, 15, V_PGI.OkDecV, '', True) + ' C au ' + DateFinExercice
  end;
(*
  if CpSens = 'C' then
  begin
    if NotePrec then
      TSoldeCre.Caption := STRFMONTANT(DbTotCP, 15, V_PGI.OkDecV, '', True)
    else
      TSoldeCre.Caption := STRFMONTANT(DbTotCE, 15, V_PGI.OkDecV, '', True);
    TSoldeDeb.Caption := '';
    TExercice.Caption   := TExercice.Caption + TSoldeCre.Caption + ' au ' + DateFinExercice
  end
  else
  begin
    if NotePrec then
      TSoldeDeb.Caption := STRFMONTANT(DbTotDP, 15, V_PGI.OkDecV, '', True)
    else
      TSoldeDeb.Caption := STRFMONTANT(DbTotDE, 15, V_PGI.OkDecV, '', True);
    TSoldeCre.Caption := '';
    TExercice.Caption   := TExercice.Caption + TSoldeDeb.Caption + ' au ' + DateFinExercice
  end;
*)

  RechercheDonnees;
  HorsFListe := False;
  FocusControle(FListe);

  // s'il y a une note de travail dans l'exercice précédent, on affiche le bouton
//  if not NotePrec then    //FQ19960
//    BZoom.Visible := YaNoteAvant;

  //si visu note précédente

  if (TFVierge(Ecran).TypeAction in [TaConsult]) or lBoPrecedent then
  begin
    BInsert.Visible       := False;
    BSupprime.Visible     := False;
    BDelete.Visible       := False;
    FBlocNote.ReadOnly    := True;
    SaisieTitre.ReadOnly  := True;
    SaisieBN.ReadOnly     := True;
    FListe.EditorMode     := False;
    FListe.Options := FListe.Options + [goRowSelect] - [goEditing] + [goTabs];
  end;

end ;

procedure TOF_CPNOTETRAVAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPNOTETRAVAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPNOTETRAVAIL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPNOTETRAVAIL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPNOTETRAVAIL.OnClose ;
begin
  FreeAndNil(FBlocNote);
  FreeAndNil(HPB);
  FreeAndNil(TobNoteTravail);
  FreeAndNil(ASaveDialog);
  Inherited ;
end ;

procedure TOF_CPNOTETRAVAIL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPNOTETRAVAIL.OnCancel () ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 13/11/2006
Modifié le ... :   /  /
Description .. : Appel des éléments de la table
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.RechercheDonnees;
var
  StrSql                : String;
  Q                     : TQuery;
  i                     : integer;

begin
  Strsql                := 'SELECT * FROM CPNOTETRAVAIL WHERE CNO_EXERCICE="' +
                           CodeExercice + '" AND CNO_GENERAL="' + CpGeneral + '" ORDER BY CNO_LIGNE';
  Q                     := OpenSQL(StrSql, True);
  TobNoteTravail.LoadDetailDB('CPNOTETRAVAIL','','',Q,False);
  Ferme(Q);

  if TobNoteTravail.Detail.Count > 1 then
  begin
    FListe.RowCount     := TobNoteTravail.Detail.Count;
    for i := 1 to TobNoteTravail.Detail.Count -1 do
    begin
      TobNoteTravail.Detail[i].AddChampSupValeur('CALCULL', 0);
      ValideLigne(i, False);
    end;
    CalculTotaux;
  end
  else
  begin
    if TobNoteTravail.Detail.Count = 0 then   // y a vraiment rien
      AjouteLigne(False, False, True); //on ajoute la ligne pour les informations
    AjouteLigne(False, True);
  end;

  StringtoRich(SaisieBN, TobNoteTravail.Detail[0].GetValue('CNO_BLOCNOTE'));
  SaisieTitre.Text      := TobNoteTravail.Detail[0].GetValue('CNO_LIBELLE');
  if SaisieTitre.Text <> '' then
  begin
    TForm(Ecran).Caption := SaisieTitre.Text;
    UpdateCaption(TForm(Ecran));
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 13/11/2006
Modifié le ... :   /  /
Description .. : Control de la ligne et mise en forme des éléments
Mots clefs ... :
*****************************************************************}
function TOF_CPNOTETRAVAIL.ValideLigne(LaLigne: integer ; SansCalc : Boolean = True ; SurSortie : Boolean = False): Boolean;
Var
  TobN                  : Tob;
  MT1                   : double;
  MT2                   : double;
  CalculL               : double;

begin
  Result                := False;

{  if SurSortie then
  begin
    if (FListe.Col = clMt1) or (FListe.Col = clMt2) then
      if not CalculCellule(FListe.Col, LaLigne) then
        exit;
  end;
}
  if SansCalc then
  begin
    if not IsValidDate(FListe.Cells[clDate, LaLigne]) then
    begin
//      PGIBox('La date saisie n''est pas valide', TForm(Ecran).Caption);
      FListe.Col        := clDate;
      FListe.Row        := LaLigne;
      exit;
    end;

    if (StrToDate(FListe.Cells[clDate, LaLigne]) > StrToDate(DateFinExercice)) or
       (StrToDate(Fliste.Cells[clDate, LaLigne]) < StrToDate(DateDebExercice)) then
    begin
      PGIInfo('La date doit être comprise dans la période de celle de l''exercice');
      FListe.Cells[clDate, LaLigne] := '';
      FListe.Col        := clDate;
      FListe.Row        := LaLigne;
      Exit;
    end;

    MT1                 := Abs(Valeur(FListe.Cells[clMt1, LaLigne]));
    MT2                 := Abs(Valeur(FListe.Cells[clMt2, LaLigne]));

    TobN                := TobNoteTravail.Detail[LaLigne];
    if TobN = nil then
      exit;

    TobN.PutValue('CNO_DATE', FListe.Cells[clDate, LaLigne]);
    TobN.PutValue('CNO_LIBELLE', FListe.Cells[clLib, LaLigne]);
    TobN.PutValue('CNO_MONTANT1', MT1);
    TobN.PutValue('CNO_MONTANT2', MT2);
    if FListe.Cells[ClSens, LaLigne] = 'Crédit' then
      TobN.PutValue('CNO_SENS', 'C')
    else
      Tobn.PutValue('CNO_SENS', 'D');

    if FListe.Cells[ClOpe, LaLigne] = '+' then
      TobN.PutValue('CNO_OPERATEUR', 'PLS')
    else
      TobN.PutValue('CNO_OPERATEUR', 'MNS');

(*
    if FBlocNote.LinesDisplayCount <> 0 then
      TobN.PutValue('CNO_BLOCNOTE', RichToString(FBlocNote))
    else
      TobN.PutValue('CNO_BLOCNOTE', '');
*)

  end
  else
  begin
    MT1                 := TobNoteTravail.Detail[LaLigne].GetValue('CNO_MONTANT1');
    MT2                 := TobNoteTravail.Detail[LaLigne].GetValue('CNO_MONTANT2');
  end;

  if Mt2 <> 0 then
    CalculL             := MT1 * MT2
  else
    CalculL             := MT1;

  TobNoteTravail.Detail[LaLigne].PutValue('CALCULL', CalculL);
  if TobNoteTravail.Detail[LaLigne].GetValue('CNO_DATE') <> iDate1900 then
    FListe.Cells[clDate, LaLigne]   := TobNoteTravail.Detail[LaLigne].GetValue('CNO_DATE');
  FListe.Cells[clLib, LaLigne]      := TobNoteTravail.Detail[LaLigne].GetValue('CNO_LIBELLE');

  if MT1 <> 0 then
    FListe.Cells[clMt1, LaLigne]    := STRFMONTANT(MT1, 15, V_PGI.OkDecV, '', True)
  else
    FListe.Cells[clMt1, LaLigne]    := '';

  if MT2 <> 0 then
    FListe.Cells[clMt2, LaLigne]    := STRFMONTANT(MT2, 15, V_PGI.OkDecV, '', True)
  else
    FListe.Cells[clMt2, LaLigne]    := '';

  if TobNoteTravail.Detail[LaLigne].GetValue('CNO_SENS') = 'C' then
    FListe.Cells[clSens, LaLigne]   := 'Crédit'
  else
    FListe.Cells[clSens, LaLigne]   := 'Débit';

  if TobNoteTravail.Detail[LaLigne].GetValue('CNO_OPERATEUR') = 'PLS' then
    FListe.Cells[clOpe, LaLigne]    := '+'
  else
    FListe.Cells[clOpe, LaLigne]    := '-';



  if TobNoteTravail.Detail[LaLigne].GetValue('CNO_SENS') = 'D' then
  begin
    FListe.Cells[clDeb, LaLigne]    := STRFMONTANT(CalculL, 15, V_PGI.OkDecV, '', True);
    FListe.Cells[clCre, LaLigne]    := '';
  end
  else
  begin
    FListe.Cells[clCre, LaLigne]    := STRFMONTANT(CalculL, 15, V_PGI.OkDecV, '', True);
    FListe.Cells[clDeb, LaLigne]    := '';
  end;

  if TobNoteTravail.Detail[LaLigne].GetValue('CNO_BLOCNOTE') <> '' then
  begin
    FListe.Cells[clBN, LaLigne]     := '#ICO#7';
//    StringtoRich(FBlocNote, TobNoteTravail.Detail[FListe.Row].GetValue('CNO_BLOCNOTE'));
  end
  else
    FListe.Cells[clBN, LaLigne]     := '';

  // on ne valide pas la ligne si le mt1 = 0
  if MT1 = 0 then
    Exit;

  Result                := True;


end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 13/11/2006
Modifié le ... :   /  /
Description .. : Ajout d'une ligne dans le tableau ou en fin
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.AjouteLigne (AuDessus : Boolean = False ; Premiere : Boolean = False ; LG0 : Boolean = False);
var
  TobN                  : Tob;
  LaLigne               : integer;
  i                     : integer;

BEGIN
  { FQ 21242 BVE 29.08.07 }
  // On ne peut pas modifier la note précédente
  if NotePrec then Exit;
  { END 21242 }
  LaLigne               := 1;
  { FQ 20963 BVE 12.07.07 }
  if HPB.Visible then
  begin
     HPB.OnClose(nil);
     HPB.Visible := false;
  end;
  { END FQ 20963 }
  if not LG0 then
  begin
    // si insertion dans le tableau
    if (FListe.Row < FListe.RowCount -1) and ( not AuDessus) then
      LaLigne           := FListe.Row +1
    else if AuDessus then
      LaLigne           := FListe.Row
      else
        LaLigne         := TobNoteTravail.Detail.Count;

    if Not Premiere then
    begin
      { FQ 20552 BVE 11.06.07 }
      // FListe.RowCount   := FListe.RowCount +1
      if not((FListe.RowCount = 2) and (TobNoteTravail.Detail.Count = 1)) then 
         FListe.InsertRow(LaLigne);
      { END FQ 20552 }
    end
    else
    begin
      FListe.RowCount   := 2;
      FListe.FixedRows  := 1;
      LaLigne           := 1;
    end;

    FListe.Options      := FListe.Options - [goRowSelect] + [goEditing] + [goTabs];
    FBlocNote.ReadOnly  := False;
    FBlocNote.Clear;
  end;

  TobN                  := Tob.Create('CPNOTETRAVAIL', TobNoteTravail, -1);
  TobN.PutValue('CNO_EXERCICE', CodeExercice);
  TobN.PutValue('CNO_GENERAL', CpGeneral);
  if LG0 then
    TobN.PutValue('CNO_LIGNE', 0)
  else
    TobN.PutValue('CNO_LIGNE', LaLigne);
//  TobN.PutValue('CNO_DATE', iDate1900);
  if StrToDate(DateFinExercice) < V_PGI.DateEntree then
    TobN.PutValue('CNO_DATE', StrToDate(DateFinExercice))
  else
    TobN.PutValue('CNO_DATE', V_PGI.DateEntree);

  TobN.PutValue('CNO_MONTANT1', 0);
  TobN.PutValue('CNO_MONTANT2', 0);
  if LG0 then
    TobN.PutValue('CNO_SENS', 'T')
  else
    TobN.PutValue('CNO_SENS', CpSens);
  TobN.PutValue('CNO_OPERATEUR', 'PLS');
  TobN.PutValue('CNO_BLOCNOTE', '');

  if LG0 then
    TobN.SetAllModifie(True);

  if not LG0 then
  begin
    TobN.AddChampSupValeur('CALCULL', 0);
    ValideLigne(LaLigne, False);
    if Premiere then
      TobN.SetAllModifie(False);

    //si on intercale dans le tableau
    if LaLigne < FListe.RowCount-1 then
    begin
      FListe.Rows[LaLigne].Clear;
      for i := LaLigne to TobNoteTravail.Detail.count -2 do
        TobNoteTravail.Detail[i].PutValue('CNO_LIGNE', i+1);
      TobNoteTravail.Detail.Sort('CNO_LIGNE');
      for i := LaLigne to FListe.RowCount-1 do
        ValideLigne(i, False);
    end;

    FListe.Col          := clDate;
    FListe.Row          := LaLigne;

  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/10/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.BinsertOnClick(Sender: TObject);
begin
  // GCO - 02/10/2007 - FQ 21553 je me sert du typeAction qui dépend de l'état
  // de la révision du dossier
  if TFVierge(Ecran).TypeAction = TaConsult then Exit;

  AjouteLigne(True);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 13/11/2006
Modifié le ... :   /  /
Description .. : Gestion des touches
Mots clefs ... :
************************************************** ***************}
procedure TOF_CPNOTETRAVAIL.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  Bloque                : boolean;
  Res                   : integer;

begin
  Bloque                := False;
  if not HorsFListe then
  begin
    Case Key of
        VK_UP : begin
                    { FQ 20643 & 20580 BVE 12.06.07 }
                    if (FListe.Col = clSens) or (FListe.Col = clOpe) or NotePrec then Exit ;
                    { END FQ 20643 & 20580 }
                  if (not ValideLigne(Fliste.Row)) and (Fliste.Row > 1) then
                  begin
//                    Fliste.Row := Fliste.Row -1;
                    key := 0;
//                    TobNoteTravail.Detail[FListe.Row +1].Free;
//                    FListe.DeleteRow(FListe.Row +1);
                      SupprimeLigne;
                  end;
                end;

        VK_DOWN : begin
                    { FQ 20643 & 20580 BVE 12.06.07 }
                    if (FListe.Col = clSens) or (FListe.Col = clOpe) or NotePrec then Exit ;
                    { END FQ 20643 & 20580 }
                    if (Fliste.Row = Fliste.RowCount -1) and (ValideLigne(Fliste.Row, True, True)) then    //FQ19959
                    begin
                      Key := 0;
                      AjouteLigne;
                    end;

                    if (Fliste.Row < Fliste.RowCount -1) and (not ValideLigne(Fliste.Row)) then
                    begin
//                      Fliste.Row := Fliste.Row + 1;
                      Key := 0;
//                      TobNoteTravail.Detail[Fliste.Row].Free;
//                      Fliste.DeleteRow(Fliste.Row -1);
                      SupprimeLigne;
                    end;
                  end;

        VK_RIGHT :  begin
                      if not NotePrec then
                      begin
                        if (FListe.Col = Fliste.ColCount -1) and (Fliste.Row < Fliste.RowCount -1) then
                        begin
                          Fliste.Col := clDate;
                          Fliste.Row := Fliste.Row + 1;
                          Key := 0;
                          Fliste.SetFocus;
                        end;
                      end;
                    end;

      VK_LEFT :  begin
                    if not NotePrec then
                    begin
                      if (Fliste.Col = clDate) and (Fliste.Row > 1) and (ValideLigne(Fliste.Row)) then
                      begin
                        key := 0;
                        Fliste.Col := Fliste.ColCount -1;
                        Fliste.Row := Fliste.Row -1;
                        Fliste.SetFocus;
                      end;
                    end;
                  end;


        VK_TAB :  begin
                    if not NotePrec then
                    begin
                      if Shift = [] then
                        if (FListe.Row = FListe.RowCount-1) and (IsValidDate(FListe.Cells[clDate, FListe.Row])) and (FListe.Col = clBN) then
                        begin
                          key := 0;
                          if ValideLigne(FListe.Row, True, True) then
                            AjouteLigne;
                        end;
                    end;

(*                    if (FListe.Col = clOpe) and (Shift = []) then
                    begin
                      FListe.Col := clDeb;
                      key := 0;
                      ComboPlusMoins.Visible := False;
                    end;
                    if (FListe.Col = clOpe) and (Shift = [ssShift]) then
                    begin
                      FListe.Col := clSens;
                      Key := 0;
                      ComboPlusMoins.Visible := False;
                    end;
*)
                  end;

        VK_INSERT : begin
                      // GCO - 02/10/2007 - FQ 21553 je me sert du typeAction qui dépend de l'état
                      // de la révision du dossier
                      if TFVierge(Ecran).TypeAction = TaConsult then Exit;

                      if not NotePrec then
                      begin
                        if (Shift = []) and (IsValidDate(FListe.Cells[clDate, FListe.Row])) then
                          AjouteLigne(True);
                      end;
                    end;

        VK_DELETE : begin
                      if not NotePrec then
                      begin
                        // GCO - 02/10/2007 - FQ 21553 je me sert du typeAction qui dépend de l'état
                        // de la révision du dossier
                        if TFVierge(Ecran).TypeAction = TaConsult then Exit;

                        if Shift = [ssCtrl] then
                          SupprimeLigne;

                        if (Shift = []) and (not FListe.EditorMode) then  //FQ19965
                          FListe.Cells[Fliste.Col, Fliste.Row]  := '';

                      end;
                    end;

        VK_F10 :  begin
                    if Shift = [] then
                      NextPrevControl(Ecran);
//                      ValideLigne(FListe.Row);
//                      CalculTotaux;
                      BValideOnClick(nil);
                  end;

        Vk_F11 :
          begin
            key         := 0;
            PopF11.Popup (Mouse.CursorPos.x, Mouse.CursorPos.y);

          end;



        VK_RETURN : begin
                    end;

        VK_F5 : begin
                  if not NotePrec then
                  begin
                    if (Shift = []) and (FListe.Col = clDate) then
                    begin
                      FListe.EditorMode := True;
                      V_PGI.ParamDateproc(FListe);
                    end;
                  end;
                end;

        VK_ESCAPE : if Shift = [] then
                      SortieSansSauvegarde;

        VK_F4 : if Shift = [ssAlt] then
                  SortieSansSauvegarde;

        VK_F7 : Begin
                  Key   := 0;
                  if not NotePrec then
                  begin
                    if (FListe.Col = clBN) or (FListe.Col = clDeb) or (FListe.Col = clCre) then
                      Bloque := True;
                    if (Shift = []) and (not Bloque) then
                      Repro(False);
                    if Shift = [ssAlt] then
                      Repro(True);
                  end;
                end;

{AO}  79 :  if Shift = [ssAlt] then
                CommMillesime;
{AB}  66 :  begin
              if Shift = [ssAlt] then
              begin
                BBlocNoteOnCLick(nil);
              end;
            end;


{AC}  67 :  if Shift = [ssAlt] then
                CommCompte;

{^P}  Ord('P') :  if Shift = [ssCtrl] then
              ImprimeTableau;

{^V}  Ord('V') :  begin
                if not NotePrec then
                begin
                  if ssCtrl in Shift then
                  begin
                    if ssShift in Shift then
                      Res := ImportClipBoard(True)
                    else
                      Res := ImportClipBoard;

                    if res <> 0 then
                      PGIInfo(strerrorclpbrd[Res], TForm(Ecran).Caption);
                  end;
                end;
              end;

    end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 13/11/2006
Modifié le ... :   /  /
Description .. : Confirme la sortie
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.SortieSansSauvegarde;
var
  Rep                   : integer;

begin
  Rep                   := mrNo;

  if not NotePrec then
    if TobNoteTravail.IsOneModifie(True) then
      Rep               := PGIAskCancel('Voulez-vous enregistrer les modifications ?', TForm(Ecran).Caption);

  Case Rep of
    mrNo                : TForm(Ecran).Close;
    mrCancel            : Exit;
    mrYes               : EnregTableau;
  end;
end;

procedure TOF_CPNOTETRAVAIL.SortieSansSauvegardeOnClick(Sender: TObject);
begin
  SortieSansSauvegarde;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : enregistrement du tableau
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.EnregTableau (Sortie : Boolean = True);
var
  StrSql                : String;
  i                     : integer;
  LaFin                 : integer;

begin

  if not NotePrec then        // n'est pas une note de l'exercice précédent
  begin
    LaFin                 := TobNoteTravail.Detail.Count-1;
    if not ValideLigne(FListe.Row) then
      if FListe.Row = FListe.RowCount -1 then
        LaFin             := TobNoteTravail.Detail.Count-2
      else
        exit;

    TobNoteTravail.Detail[0].PutValue('CNO_BLOCNOTE', RichToString(SaisieBN));
    TobNoteTravail.Detail[0].PutValue('CNO_LIBELLE', SaisieTitre.Text);

    Try
      BEGINTRANS;
      StrSql              := 'DELETE CPNOTETRAVAIL WHERE CNO_EXERCICE="' +
                             CodeExercice+'" AND CNO_GENERAL="' + CpGeneral + '"';

      ExecuteSQL(StrSql);
      for i := 0 to LaFin do
      begin
        TobNoteTravail.Detail[i].SetAllModifie(True);
        TobNoteTravail.Detail[i].InsertOrUpdateDB(False);
      end;
      COMMITTRANS;
    Except
      on E: Exception do
      begin
        ROLLBACK;
      end;
    end;
  end;
  if Sortie then
    TForm(Ecran).Close;
end;

procedure TOF_CPNOTETRAVAIL.BValideOnClick(Sender: TObject);
begin
  EnregTableau;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Traitement en entrant dans une cellule
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.FListeCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
(*
var
  r                     : TRect ;
*)

begin
(*
  if  (FListe.Col = clOpe ) and ( FListe.Row > 0 ) and (not NotePrec) then
  begin
    r                          := FListe.CellRect(clOpe,Fliste.Row);
    ComboPlusMoins.Top         := r.Top+FListe.top ;
    ComboPlusMoins.Left        := r.Left+Fliste.Left;
    ComboPlusMoins.Width       := r.Right-r.Left ;
    ComboPlusMoins.Height      := r.Bottom-r.Top ;
    ComboPlusMoins.ComboWidth  := 10;
    ComboPlusMoins.Visible     := True ;
    ComboPlusMoins.Value       := TobNoteTravail.Detail[FListe.Row].GetValue('CNO_OPERATEUR');
    ComboPlusMoins.SetFocus;

  end
  else
  begin
     ComboPlusMoins.Visible     := False ;
  end;
  *)
end;

(*
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 21/11/2006
Modifié le ... :   /  /
Description .. : sur une nouvelle valeur de la combo plusmoins
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.ComboChanged (Sender : Tobject);
begin
   FListe.Cells[clOpe, FListe.Row] := ComboPlusMoins.Text;
end;
*)

(*
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 21/11/2006
Modifié le ... :   /  /
Description .. : sur action d'une touche dans la combo
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.ComboKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_RIGHT :  begin
                  if Shift = [] then
                  begin
                    key := 0;
                    ComboPlusMoins.Visible := False;
                    FListe.SetFocus;
                    FListe.Col := clDeb;
                  end;
                end;

    VK_LEFT :   begin
                  if Shift = [] then
                  begin
                    key := 0;
                    ComboPlusMoins.Visible := False;
                    FListe.SetFocus;
                    FListe.Col := clSens;
                  end;
                end;

    VK_F4 :   begin
                ComboPlusMoins.DroppedDown  := True;
              end;

    107 :     begin
                ComboPlusMoins.Text := '+';
              end;
    109 :     begin
                ComboPlusMoins.Text := '-';
              end;
  end;

end;
*)


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /    
Description .. : Traitement en sortie de cellule
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  val : string;
begin
//  Inherited ;
  if ACol = clLib then
  begin
    if Length(FListe.Cells[clLib, ARow]) > 35 then
      FListe.Cells[clLib, ARow] := Copy(FListe.Cells[clLib, ARow], 1, 35);
  end;
  { FQ 19993 BVE 15.06.07 }
  if ( ACol = clDate ) then
  begin
     val := FListe.Cells[ACol,ARow];
     if IsNumeric(val) then
     begin
        if Length(val) = 8 then
        begin
           FListe.Cells[ACol,ARow] := Copy(val,1,2) + '/' + Copy(val,3,2) + '/' + Copy(val,5,4);
        end;
     end; 
  end;                                                                                          
  { END FQ 19993 }
(*
  if not ValideLigne(ARow, True, True) then   // force la valide avec sortie de cellule
  begin
    FListe.Col          := ACol;
    FListe.Row          := ARow;
  end
  else
*)
    ValideLigne(ARow, True, True);
    CalculTotaux;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Calcul des totaux
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.CalculTotaux;
var
  TotDeb                : double;
  TotCre                : double;
  RapDeb                : double;
  RapCre                : double;
  i                     : integer;

begin
  TotDeb                := 0;
  TotCre                := 0;
  RapDeb                := 0;
  RapCre                := 0;

  for i := 1 to TobNoteTravail.Detail.Count -1 do
  begin
    if TobNoteTravail.Detail[i].GetValue('CNO_SENS') = 'D' then
    begin
      if TobNoteTravail.Detail[i].GetValue('CNO_OPERATEUR') = 'PLS' then
        TotDeb          := TotDeb + TobNoteTravail.Detail[i].GetValue('CALCULL')
      else
        TotDeb          := TotDeb - TobNoteTravail.Detail[i].GetValue('CALCULL');
    end
    else
    begin
      if TobNoteTravail.Detail[i].GetValue('CNO_OPERATEUR') = 'PLS' then
        TotCre          := TotCre + TobNoteTravail.Detail[i].GetValue('CALCULL')
      else
        TotCre          := TotCre - TobNoteTravail.Detail[i].GetValue('CALCULL');
    end;
  end;

  // Affichage du solde
  TBalCre.Caption   := '';
  TBalDeb.Caption   := '';
  if (TotDeb - TotCre) > 0 then
    TBalDeb.Caption   := STRFMONTANT(Abs((TotDeb - TotCre)), 15, V_PGI.OkDecV, '', True)
  else  if (TotDeb - TotCre) < 0 then
    TBalCre.Caption   := STRFMONTANT(Abs((TotDeb - TotCre)), 15, V_PGI.OkDecV, '', True);

//Calcul du pourcentage par rapport au solde
  if (CpSens = 'D') or (CpSens = 'M') then
  begin
//    if VALEUR(TSoldeDeb.Caption) <> 0 then      //FQ20016
    if VALEUR(TSoldeDeb.Caption) > 0 then
    begin
      RapDeb            := (TotDeb - TotCre) / VALEUR(TSoldeDeb.Caption) * 100;
    end;
(*
    if Assigned(TBalDeb) then
    begin
      TBalDeb.Caption   := STRFMONTANT(Abs((TotDeb - TotCre)), 15, V_PGI.OkDecV, '', True);
      TBalCre.Caption   := '';
    end;
*)
  end;

  if (CpSens = 'C')  then
  begin
//    if VALEUR(TSoldeCre.Caption) <> 0 then    //FQ20016
    if VALEUR(TSoldeCre.Caption) > 0 then
    begin
      RapCre            := (TotCre - TotDeb) / VALEUR(TSoldeCre.Caption) * 100;
    end;
(*
    if Assigned(TbalCre) then
    begin
      TBalCre.Caption   := STRFMONTANT(Abs((TotCre - TotDeb)), 15, V_PGI.OkDecV, '', True);
      TBalDeb.Caption   := '';
    end;
*)
  end;

  if TotDeb <> 0 then
    TTotDeb.Caption     := STRFMONTANT(TotDeb, 15, V_PGI.OkDecV, '', True)
  else
    TTotDeb.Caption     := '';
  if TotCre <> 0 then
    TTotCre.Caption     := STRFMONTANT(TotCre, 15, V_PGI.OkDecV, '', True)
  else
    TTotCre.Caption     := '';
  { FQ 20016 BVE 13.06.07 }
  if RapDeb > 0 then
    TPDeb.Caption       := STRFMONTANT(RapDeb, 15, V_PGI.OkDecV, '%', True)
  else if RapDeb < 0 then
    TPDeb.Caption       := 'NS'
  else
    TPDeb.Caption       := '';
  if RapCre > 0 then
    TPCre.Caption       := STRFMONTANT(RapCre, 15, V_PGI.OkDecV, '%', True)
  else if RapCre < 0 then
    TPCre.Caption       := 'NS'
  else
    TPCre.Caption       := '';
  { END FQ 20016 }


end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Suppression de la ligne en cours
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.SupprimeLigne;
var i : integer;
begin
  if TobNoteTravail.Detail.Count <= 1 then exit;

  { FQ 21242 BVE 29.08.07 }
  // On ne peut pas modifier la note précédente
  if NotePrec then Exit;
  { END 21242 }

  if PGIAsk('Voulez vous supprimer la ligne ?', TForm(Ecran).Caption) = mrYes then
  begin
    TobNoteTravail.Detail[FListe.Row].Free;
    FListe.DeleteRow(FListe.Row);
    CalculTotaux;
//    if FListe.RowCount = 2 then     //FQ19967

    { FQ 19967 BVE 13.06.07
    if TobNoteTravail.Detail.Count = 0 then }
    if TobNoteTravail.Detail.Count = 1 then
    { END FQ 19967 }
    begin
      AjouteLigne(False, True);
    end;

    for i := 1 to TobNoteTravail.Detail.Count -1 do
    begin
      TobNoteTravail.Detail[i].PutValue('CNO_LIGNE', i);
    end;
    FBLocNote.Clear;
    if TobNoteTravail.Detail.Count -1 > 1 then    // iln'y a que le descriptif général
      StringtoRich(FBlocNote, TobNoteTravail.Detail[FListe.Row].GetValue('CNO_BLOCNOTE'));
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.BSupprimeLigneOnClick(Sender: TObject);
begin
  // GCO - 02/10/2007 - FQ 21553 je me sert du typeAction qui dépend de l'état
  // de la révision du dossier
  if TFVierge(Ecran).TypeAction = TaConsult then Exit;

  SupprimeLigne;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Duplication de la lgibe précédente
Suite ........ : - RTotal = True -> recopie la ligne précédente
Suite ........ : - RTotal = False -> recopie le champ au dessus
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.Repro(RTotal: Boolean);
var
  LaLigne               : integer;
  LaCol                 : integer;

begin
  { FQ 21242 BVE 29.08.07 }
  // On ne peut pas modifier la note précédente
  if NotePrec then Exit;
  { END 21242 }

  LaLigne               := FListe.Row;
  if LaLigne = 1 then
    exit;

  if RTotal then
  begin
    TobNoteTravail.Detail[LaLigne].PutValue('CNO_DATE', FListe.Cells[clDate, LaLigne-1]);
    TobNoteTravail.Detail[LaLigne].PutValue('CNO_LIBELLE', FListe.Cells[clLib, LaLigne-1]);
    TobNoteTravail.Detail[LaLigne].PutValue('CNO_MONTANT1', Valeur(FListe.Cells[clMt1, LaLigne-1]));
    TobNoteTravail.Detail[LaLigne].PutValue('CNO_MONTANT2', Valeur(FListe.Cells[clMt2, LaLigne-1]));
    if FListe.Cells[clSens, LaLigne-1] = 'Crédit' then
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_SENS', 'C')
    else
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_SENS', 'D');
    if FListe.Cells[clOpe, LaLigne-1] = '+' then
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_OPERATEUR', 'PLS')
    else
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_OPERATEUR', 'MNS');
    TobNoteTravail.Detail[LaLigne].PutValue('CNO_BLOCNOTE', TobNoteTravail.Detail[LaLigne-1].GetValue('CNO_BLOCNOTE'));
    ValideLigne(LaLigne, False);
    CalculTotaux;
  end
  else
  begin
    LaCol               := FListe.Col;

    if LaCol = clDate then
    Begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_DATE', TobNoteTravail.Detail[LaLigne-1].GetValue('CNO_DATE'));
      ValideLigne(LaLigne, False);
    end;
    if LaCol = clLib then
    begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_LIBELLE', TobNoteTravail.Detail[LaLigne-1].GetValue('CNO_LIBELLE'));
      ValideLigne(LaLigne, False);
    end;
    if LaCol = clMt1 then
    begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_MONTANT1', TobNoteTravail.Detail[LaLigne-1].GetValue('CNO_MONTANT1'));
      ValideLigne(LaLigne, False);
      CalculTotaux;
    end;
    if LaCol = clMt2 then
    begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_MONTANT2', TobNoteTravail.Detail[LaLigne-1].GetValue('CNO_MONTANT2'));
      ValideLigne(LaLigne, False);
      CalculTotaux;
    end;
    if LaCol = clsens then
    begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_SENS', TobNoteTravail.Detail[LaLigne-1].GetValue('CNO_SENS'));
      ValideLigne(LaLigne, False);
    end;
    if LaCol = clOpe then
    begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_OPERATEUR', TobNoteTravail.Detail[LaLigne-1].GetValue('CNO_OPERATEUR'));
      ValideLigne(LaLigne, False);
    end;
    if LaCol = clBN then
    begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_BLOCNOTE', TobNoteTravail.Detail[LaLigne-1].GetValue('CNO_BLOCNOTE'));
      ValideLigne(LaLigne, False);
    end;
  end;


end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Commentaire Millésimé
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.CommMillesime;
begin
  CPLanceFiche_CPREVDOCTRAVAUX(CpGeneral, '', CodeExercice, 3)
end;

procedure TOF_CPNOTETRAVAIL.CommMillesimeOnClick(Sender: Tobject);
begin
  CommMillesime;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... : 30/05/2006
Description .. : Commentaire du compte général
Suite ........ : GCO - 30/05/2006 - FQ 20446
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.CommCompte;
begin
  CPLanceFiche_CPREVDOCTRAVAUX(CpGeneral, '', CodeExercice, 4)
end;

procedure TOF_CPNOTETRAVAIL.CommcompteOnClick(Sender: Tobject);
begin
  CommCompte;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Detruit le tableau et la table associée
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.BDeleteOnCLick(Sender: Tobject);
var
  StrSql                : String;
  i                     : integer;

begin
  { FQ 21242 BVE 29.08.07 }
  // On ne peut pas modifier la note précédente
  if NotePrec then Exit;
  { END 21242 }

  if TobNoteTravail.Detail.Count = 0 then
    exit;

  if PGIAsk('Voulez vous vraiment détruire ce tableau ?', TForm(Ecran).Caption) = mrYes then
  begin
    StrSql              := 'DELETE CPNOTETRAVAIL WHERE CNO_EXERCICE="' +
                           CodeExercice + '" AND CNO_GENERAL="' + CpGeneral + '"';

    ExecuteSQL(StrSql);
    TobNoteTravail.Free;
    SaisieTitre.Text    := '';
    SaisieBN.Clear;
    TobNoteTravail      := Tob.Create('CPNOTETRAVAIL', nil, -1);
    for i := FListe.RowCount -1  downto 1 do
    begin
      FListe.DeleteRow(i);
    end;
//    CalculTotaux;
//    AjouteLigne(False, False, True);   //ajout de la ligne informations
//    AjouteLigne(False, True);
    TForm(Ecran).Close;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 14/11/2006
Modifié le ... :   /  /
Description .. : Gestion du bloc note
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.BBlocNoteOnCLick(Sender: Tobject);
var
  LaLigne               : integer;

begin
  if Sender = nil then         // si appel depuis alt+B
  begin
    if HPB.visible = False then
    begin
      FBLocNote.Clear;
      StringtoRich(FBlocNote, TobNoteTravail.Detail[FListe.Row].GetValue('CNO_BLOCNOTE'));
      HPB.Visible       := True;
      BBlocNote.Down    := True;
//      FBlocNote.SetFocus;
    end
    else
    begin
      HPB.Visible       := False;
      BBlocNote.Down    := False;
      FListe.SetFocus;
    end;
  end
  else                         // appuie sur bouton
    begin
    If BBlocNote.Down then
    begin
      FBLocNote.Clear;
      StringtoRich(FBlocNote, TobNoteTravail.Detail[FListe.Row].GetValue('CNO_BLOCNOTE'));
      HPB.Visible       := True;
//      FBlocNote.SetFocus;
    end
    else
    begin
      HPB.Visible       := False;
      FListe.SetFocus;
    end;
  end;

  if not BBlocNote.Down then
  begin
    LaLigne             := FListe.Row;
    if FBlocNote.LinesDisplayCount <> 0 then
    begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_BLOCNOTE', RichToString(FBlocNote));
      FListe.Cells[clBN, LaLigne] := '#ICO#7';
    end
    else
    begin
      TobNoteTravail.Detail[LaLigne].PutValue('CNO_BLOCNOTE', '');
      FListe.Cells[clBN, LaLigne] := '';
    end;
  end;
end;

procedure TOF_CPNOTETRAVAIL.HPBOnClose(Sender: Tobject);
var
  LaLigne               : integer;

begin
  LaLigne               := FListe.Row;
  TobNoteTravail.Detail[LaLigne].PutValue('CNO_BLOCNOTE', RichToString(FBlocNote));
  if FBlocNote.LinesDisplayCount <> 0 then
    FListe.Cells[clBN, LaLigne] := '#ICO#7'
  else
    FListe.Cells[clBN, LaLigne] := '';
  BBlocNote.Down        := False;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/11/2006
Modifié le ... :   /  /    
Description .. : sur dépacement dans le tableau
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.FListeOnRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  HPB.Visible := BBlocNote.Down;
  FBLocNote.Clear;
  StringtoRich(FBlocNote, TobNoteTravail.Detail[FListe.Row].GetValue('CNO_BLOCNOTE'));
end;

procedure TOF_CPNOTETRAVAIL.FlisteOnRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  MajBlocNote(Ou);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/11/2006
Modifié le ... :   /  /    
Description .. : saisie du bloc note général
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.SaisieBNOnExit(Sender: TObject);
begin
  TobNoteTravail.Detail[0].PutValue('CNO_BLOCNOTE', RichToString(SaisieBN));
  HorsFListe            := False;
end;

procedure TOF_CPNOTETRAVAIL.SaisieBNOnEnter(Sender: Tobject);
begin
  HorsFListe            := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/11/2006
Modifié le ... :   /  /
Description .. : saisie du titre de la note
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.SaisieTitreOnExit(Sender: TObject);
begin
  TobNoteTravail.Detail[0].PutValue('CNO_LIBELLE', SaisieTitre.Text);
  if SaisieTitre.Text <> '' then
  begin
    TForm(Ecran).Caption := SaisieTitre.Text;
    UpdateCaption(TForm(Ecran));
  end
  else
    TForm(Ecran).Caption := 'Note de travail : ' + CpGeneral + ' - ' + StrLibCpt;
  HorsFListe            := False;

end;

procedure TOF_CPNOTETRAVAIL.SaisieTitreOnEnter(Sender: Tobject);
begin
  HorsFListe := True;
end;

(*
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 17/11/2006
Modifié le ... :   /  /
Description .. : reprise code de LG
Suite ........ : calcul dans cellule
Mots clefs ... : 
*****************************************************************}
function TOF_CPNOTETRAVAIL.GetFormulePourCalc(Formule : hstring) : Variant ;
begin
  Result:=#0 ;
  Formule:=AnsiUpperCase(Trim(Formule));
  if Formule = '' then
    Exit ;
end ;
*)

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/11/2006
Modifié le ... :   /  /    
Description .. : export du tableau vers différents formats
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.BExportOnClick(Sender: TObject);
Var
  SauveHint             : String;

begin
  if ASaveDialog.Execute then
  begin
    if ASaveDialog.FilterIndex = 5 then //html
    begin
      SauveHint         := FListe.Hint;
      FListe.Hint       := Ecran.Caption;
      ExportGrid(FListe, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
      FListe.Hint       := SauveHint;
    end
    else
      ExportGrid(FListe, nil, ASaveDialog.FileName, ASaveDialog.FilterIndex, True);
  end;

end;


           (*
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/11/2006
Modifié le ... : 17/11/2006
Description .. : fonction pour faire un calcul dans la cellule
Suite ........ : reprise du code de LG
Mots clefs ... : 
*****************************************************************}
function TOF_CPNOTETRAVAIL.CalculCellule(LaCol, LaLigne: integer): Boolean;
begin
  if (pos('+',FListe.Cells[LaCol, LaLigne])>0) or (pos('-',FListe.Cells[LaCol, LaLigne])>0) or (pos('/',FListe.Cells[LaCol, LaLigne])>0) or (pos('*',FListe.Cells[LaCol, LaLigne])>0) then
  begin
    if (not VH^.MontantNegatif) and ( pos( Copy(FListe.Cells[LaCol, LaLigne], 1, 1), '+;-;/;*' ) > 0 ) then
    begin
      FListe.Cells[LaCol, LaLigne]:= GFormule(FListe.Cells[LaCol, LaLigne]+'}', GetFormulePourCalc, nil, 1);
    end
    else
    begin
    if Pos(FListe.Cells[LaCol, LaLigne][1],'/')>-1 then
      FListe.Cells[LaCol, LaLigne]:='{"#.###,0"'+FListe.Cells[LaCol, LaLigne]+'}';
      FListe.Cells[LaCol, LaLigne]:=GFormule('{'+FListe.Cells[LaCol, LaLigne]+'}', GetFormulePourCalc, nil, 1)
    end;
  end;
  Result := True;
end;
         *)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/07/2007
Modifié le ... :   /  /    
Description .. : FQ 21068
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.OnClickCollerEcriture(Sender: TObject);
begin
  ImportClipBoard( False );
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 17/11/2006
Modifié le ... :   /  /
Description .. : controle et gestion du presse papier depuis un ctrl-c de la
Suite ........ : consultation des ecritures.
Suite ........ : on rapporte 1 ou plusieures ecritures dans le tableau
Mots clefs ... :
*****************************************************************}
function TOF_CPNOTETRAVAIL.ImportClipBoard (EstCumul : Boolean) : integer;
Var
  StrClpbrd             : WideString;
  Ecriture              : String;
  Element               : string;
  Ipos                  : integer;
  TobN                  : tob;
  Index                 : integer;
  AncI                  : integer;
  i                     : integer;
  NbL                   : integer;

begin
  Result := 0;
  { FQ 21242 BVE 29.08.07 }
  // On ne peut pas modifier la note précédente
  if NotePrec then Exit;
  { END 21242 }

  //est-ce du texte ?
  if not Clipboard.HasFormat(CF_TEXT) then
    Result              := 1
  else
  begin
    Index               := FListe.Row;
    AncI                := FListe.Row;
    NbL                 := 0;
    // BVE 11.06.07
    TobN                := nil;

    StrClpbrd := Clipboard.AsText;

    { FQ 21273 BVE 28.08.07 }
    // On recherche ;
    if (Pos('";"',StrClpbrd) = 0) then
    begin        
      Result            := 1;
      Exit;
    end;
    { END FQ 21273 }

    if (Trim(StrClpbrd) = '') then
    begin
      Result            := 2;
      Exit;
    end;
    Ipos                := Pos(chr(13)+chr(10), StrClpbrd);
    Ecriture            := Copy(StrClpbrd, 1, Ipos-1);
    StrClpbrd           := Copy(StrClpbrd,Ipos+3, Length(StrClpbrd));

    while ecriture <> '' do
    begin
      Element           := StringReplace_(READTOKENST(Ecriture), '"', '',[rfReplaceAll	]);
      while Element <> '' do
      begin
        if (Not EstCumul) or ((EstCumul) and (Nbl = 0)) then
          TobN          := Tob.Create('CPNOTETRAVAIL', TobNoteTravail, -1);
        if not (Element = CodeExercice) then
        begin
          TobN.Free;
          Result        := 3;
          exit;
        end;
        TobN.PutValue('CNO_EXERCICE', Element);

        Element         := StringReplace_(READTOKENST(Ecriture), '"', '',[rfReplaceAll	]);
        if not (Element = CpGeneral) then
        begin
          TobN.Free;
          Result        := 4;
          Exit;
        end;
        TobN.PutValue('CNO_GENERAL', Element);


        Element         := StringReplace_(READTOKENST(Ecriture), '"', '',[rfReplaceAll	]);
        if (EstCumul) and (NbL = 0) then
        begin
          TobN.PutValue('CNO_LIGNE', Index);
          TobN.PutValue('CNO_DATE', Element);
        end
        else if not EstCumul then
        begin
          TobN.PutValue('CNO_LIGNE', Index);
          TobN.PutValue('CNO_DATE', Element);
        end;

        Element         := StringReplace_(READTOKENST(Ecriture), '"', '',[rfReplaceAll	]);
        if (EstCumul) and (Nbl = 0) then
          TobN.PutValue('CNO_LIBELLE', 'Cumul ' + Element)
        else if not EstCumul then
          TobN.PutValue('CNO_LIBELLE', Element);


        Element         := StringReplace_(READTOKENST(Ecriture), '"', '',[rfReplaceAll	]);
        if (EstCumul) and (NbL > 0) then
          TobN.PutValue('CNO_MONTANT1', Valeur(Element) + TobN.GetValue('CNO_MONTANT1'))
        else if (not EstCumul) or ((EstCumul) and (NbL = 0)) then
          TobN.PutValue('CNO_MONTANT1', Valeur(Element));

        TobN.PutValue('CNO_MONTANT2', 0);

        Element         := StringReplace_(READTOKENST(Ecriture), '"', '',[rfReplaceAll	]);
        if not((Element = 'C') or (Element = 'D')) then
        begin
          TobN.Free;
          Result        := 5;
          Exit;
        end;
        TobN.PutValue('CNO_SENS', Element);
        TobN.PutValue('CNO_OPERATEUR', 'PLS');
        TobN.PutValue('CNO_BLOCNOTE', '');
        TobN.AddChampSupValeur('CALCULL', 0);

        //normalement, il y a plus rien. Dans le cas contraire on annule tout
        Element         := StringReplace_(READTOKENST(Ecriture), '"', '',[rfReplaceAll	]);
        if Trim(Element) <> '' then
        begin
          Result        := 6;
          TobN.Free;
          Exit;
        end;
      end;

      Inc (NbL);
      Inc (Index);

      Ipos              := Pos(chr(13)+chr(10), StrClpbrd);
      Ecriture          := Copy(StrClpbrd, 1, Ipos-1);
      StrClpbrd         := Copy(StrClpbrd,Ipos+3, Length(StrClpbrd));

    end;

    for i := AncI to FListe.RowCount-1 do
      TobNoteTravail.Detail[i].PutValue('CNO_LIGNE', i+index);

    FListe.RowCount     := TobNoteTravail.Detail.Count;

    TobNoteTravail.Detail.Sort('CNO_LIGNE');
    for i := 1 to TobNoteTravail.Detail.Count -1 do
    begin
      ValideLigne(i, False);
    end;

  end;
  CalculTotaux;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 17/11/2006
Modifié le ... :   /  /
Description .. : impression du tableau
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.ImprimeTableau;
begin
  if not NotePrec then
    EnregTableau(False);
  TobNoteTravail.Modifie  := False;

  TRic.CPLanceEtat_NoteTravail( CPGeneral, CodeExercice, 0 );
end;

procedure TOF_CPNOTETRAVAIL.BImprimeTableauOnCLick(Sender: TObject);
begin
  ImprimeTableau;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/11/2006
Modifié le ... :   /  /
Description .. : passage en mode visualisation d'une note de travail d'un
Suite ........ : exercice précédent
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.BZoomOnClick(Sender: Tobject);
var
  Lexe                  : String;

begin
  with Sender as TMenuItem do
  begin
    Lexe                :=Name;
    if Lexe = 'ANC' then
      CPLanceFiche_CPNOTETRAVAIL( CpGeneral, True, VH^.Precedent.Code)
    else
    begin
      Lexe              := Copy(Lexe, 3, length(Lexe));
      CPLanceFiche_CPNOTETRAVAIL( CpGeneral, True, Lexe);
    end;
  end;
//  VisuAncienneNote;       //FQ19960
end;

{                    //FQ19960
procedure TOF_CPNOTETRAVAIL.VisuAncienneNote;
//var
  //StrSql        : String;
  //StrArg        : String;
  //FinExerPrec   : String;
  //ExercicePrec  : String;
  //Q             : TQuery;
begin
  //FinExerPrec := USDATETIME(StrToDate(DateDebExercice)-1);
  //StrSql := 'SELECT EX_EXERCICE FROM EXERCICE WHERE EX_DATEFIN = "'+FinExerPrec+'"';
  //Q := OpenSQL(StrSql, True);
  //ExercicePrec := Q.Fields[0].AsString;
  //Ferme(Q);

  //StrArg := 'COMPTE='+CpGeneral+';LIBELLE='+StrLibCpt+';EXERCICE='+ExercicePrec+
  //          ';SENS='+CpSens+';CREDITE=0;DEBITE=0'+';CREDIP='+FloatToStr(DbTotCP)+
  //          ';DEBITP='+FloatToStr(DbTotDP)+';PRECEDENT=OUI';

  CPLanceFiche_CPNoteTravail(CPGeneral, True);
end;
}

function TOF_CPNOTETRAVAIL.YaNoteAvant: boolean;
var
  StrSql                : String;
//  FinExerPrec           : String;   //FQ19960

begin
//  FinExerPrec           := USDATETIME(StrToDate(DateDebExercice)-1);
  StrSql                := 'SELECT CNO_EXERCICE FROM CPNOTETRAVAIL WHERE CNO_GENERAL = "' + CpGeneral +
                           '" AND CNO_EXERCICE = "' + VH^.Precedent.Code + '"';
  if ExisteSQL(StrSql) then
    Result              := True
  else
    Result              := False;

end;

procedure TOF_CPNOTETRAVAIL.BlocNoteOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case key of

      66 :  begin
              if Shift = [ssAlt] then
              begin
                BBlocNoteOnCLick(nil);
              end;
            end;
  end;
end;


procedure TOF_CPNOTETRAVAIL.ReporElonClick(Sender: Tobject);
begin
  if not NotePrec then
  begin
    if (FListe.Col <> clBN) and (FListe.Col <> clDeb) and (FListe.Col <> clCre) then
      Repro(False);
  end;

end;

procedure TOF_CPNOTETRAVAIL.ReproTotOnClick(Sender: Tobject);
begin
  if not NotePrec then
    Repro(True);

end;

procedure TOF_CPNOTETRAVAIL.ObsNoteOnClick(Sender: Tobject);
begin
  PGNote.ActivePage             := PGNote.Pages[1];
  SetFocusControl('SAISIEBN');
end;

procedure TOF_CPNOTETRAVAIL.RenomerNoteOnClick(Sender: Tobject);
begin
  PGNote.ActivePage             := PGNote.Pages[1];
  SetFocusControl('SAISIETITRE');
end;


procedure TOF_CPNOTETRAVAIL.BlocnoteOnclickF11(Sender: Tobject);
begin
  FBLocNote.Clear;
  StringtoRich(FBlocNote, TobNoteTravail.Detail[FListe.Row].GetValue('CNO_BLOCNOTE'));
  HPB.Visible           := True;
  BBlocNote.Down        := True;
//  FBlocNote.SetFocus;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : tri de la note de travail selon le type sélectionné
Suite ........ : Gestion d'un booléen lié pour changer le sens du tri
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNOTETRAVAIL.LeTri(LeType: String);
var
  LaTobTri              : Tob;
  TobT                  : Tob;
  TobS                  : Tob;
  LeChamp               : String;
  i                     : integer;

begin
  //on recupère la valeur du blocnote
  MajBlocNote(Fliste.Row);

  LaTobTri              := Tob.Create('tob detri', nil, -1);
  LaTobTri.Dupliquer(TobNoteTravail, True, True);
  TobT                  := LaTobTri.FindFirst(['CNO_LIGNE'],[0],False);
  TobT.Free;

  if LaTobTri.Detail.count <1 then
    exit;

  if LeType = 'DATE' then
  begin
    TriLibDesc          := True;
    TriMtDesc           := True;
    LeChamp             := 'CNO_DATE';
    if not TriDateDesc then
    begin
      LeChamp           := '-' + LeChamp;
      TriDateDesc       := True;
    end
    else
      TriDateDesc       := False;
  end

  else if LeType = 'LIB' then
  begin
    TriDateDesc         := True;
    TriMtDesc           := True;
    LeChamp             := 'CNO_LIBELLE';
    if not TriLibDesc then
    begin
      LeChamp           := '-' + LeChamp;
      TriLibDesc        := True;
    end
    else
      TriLibDesc        := False;
  end
  else if LeType = 'MT' then
  begin
    TriDateDesc         := True;
    TriLibDesc          := True;
    LeChamp             := 'CNO_MONTANT1';
    if not TriMtDesc then
    begin
      LeChamp           := '-' + LeChamp;
      TriMtDesc         := True;
    end
    else
      TriMtDesc         := False;
  end;

  LaTobTri.Detail.Sort(LeChamp);


  for i := 0 to LaTobTri.Detail.Count -1 do
  begin
    TobS                := TobNoteTravail.Detail[i+1];
    TobT                := LaTobTri.Detail[i];
    tobS.PutValue('CNO_DATE', TobT.GetValue('CNO_DATE'));
    tobS.PutValue('CNO_LIBELLE', TobT.GetValue('CNO_LIBELLE'));
    tobS.PutValue('CNO_MONTANT1', TobT.GetValue('CNO_MONTANT1'));
    tobS.PutValue('CNO_MONTANT2', TobT.GetValue('CNO_MONTANT2'));
    tobS.PutValue('CNO_SENS', TobT.GetValue('CNO_SENS'));
    tobS.PutValue('CNO_OPERATEUR', TobT.GetValue('CNO_OPERATEUR'));
    tobS.PutValue('CNO_BLOCNOTE', TobT.GetValue('CNO_BLOCNOTE'));
    ValideLigne(i+1, False);
  end;
  CalculTotaux;

  // et surtout ne pas oublier de rafraichir le blocnote ...
  StringtoRich(FBlocNote, TobNoteTravail.Detail[FListe.Row].GetValue('CNO_BLOCNOTE'));
end;

procedure TOF_CPNOTETRAVAIL.TridateOnClick(Sender: Tobject);
begin
  LeTri('DATE');
end;

procedure TOF_CPNOTETRAVAIL.TriLibelleOnClick(Sender: tobject);
begin
  LeTri('LIB');
end;

procedure TOF_CPNOTETRAVAIL.TriMontantOnClick(Sender: Tobject);
begin
  LeTri('MT');
end;

procedure TOF_CPNOTETRAVAIL.CycleCommentaireOnClick(Sender: Tobject);
begin
  CPLanceFiche_CPRevDocTravaux( CpGeneral, '', CodeExercice, 0 );
end;

procedure TOF_CPNOTETRAVAIL.CycleObjectifOnClick(Sender: Tobject);
begin
  CPLanceFiche_CPRevDocTravaux( CpGeneral, '', CodeExercice, 1 );
end;

procedure TOF_CPNOTETRAVAIL.CycleSyntheseOnClick(Sender: Tobject);
begin
  CPLanceFiche_CPRevDocTravaux( CpGeneral, '', CodeExercice, 2 );
end;

procedure TOF_CPNOTETRAVAIL.DeplaceBlocNote;
var
  LeX                   : integer;
  LeY                   : integer;
  LeW                   : integer;
  LeH                   : integer;
  LePanel               : THPanel;
  LePanel2              : THPanel;
  LeLabel               : THLabel;
  LeFD11                : THLabel;

begin
  LePanel               := THPanel(GetControl('FE_TOTAUX'));
  LePanel2              := THPanel(GetControl('FE_INFOS'));
  LeLabel               := THLabel(GetControl('TSOLDE'));
  LeFD11                := THLabel(GetControl('FD11'));

  LeX                   := LePanel.Left + LeLabel.Left + LeLabel.Width + 10;
  LeY                   := LePanel.Top + LePanel2.Height + 5;
  LeW                   := LeFD11.Left - LeLabel.Width - LeLabel.Left - 10;
  LeH                   := LePanel.Height - 10;

  HPB.Left              := LeX + TForm(Ecran).Left;
  HPB.Top               := LeY + TForm(Ecran).Top;
  HPB.Width             := LeW;
  HPB.Height            := LeH;


end;

procedure TOF_CPNOTETRAVAIL.HPBOnVisibleChanged(Sender: Tobject);
begin
  if HPB.Visible = True then
    DeplaceBlocNote;
end;

procedure TOF_CPNOTETRAVAIL.PgNoteOnChanged(Sender: Tobject);
begin
  if (PGNote.ActivePageIndex <> 0) and (HPB.Visible = True) then
  begin
    HPB.Visible       := False;
  end;
  if (PGNote.ActivePageIndex = 0) and (BBlocNote.down = True) then
  begin
    HPB.Visible       := True;
  end;
end;


procedure TOF_CPNOTETRAVAIL.MajBlocNote(LaLigne: integer);
begin
  { FQ 21242 BVE 29.08.07 }
  // On ne peut pas modifier la note précédente
  if NotePrec then Exit;
  { END 21242 }

    //quoi qu'il se passe on met à jour la tob pour les blocnotes
  if FBlocNote.LinesDisplayCount <> 0 then
    TobNoteTravail.Detail[LaLigne].PutValue('CNO_BLOCNOTE', RichToString(FBlocNote))
  else
    TobNoteTravail.Detail[LaLigne].PutValue('CNO_BLOCNOTE', '');


end;

Initialization
  registerclasses ( [ TOF_CPNOTETRAVAIL ] ) ;
end.
