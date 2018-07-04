{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ZECRIMVT ()
Mots clefs ... : TOF;ZECRIMVT
*****************************************************************}

Unit ZECRIMVT_TOF;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    MaineAGL,
    eMul,
    UtileAGL,
{$ELSE}
    FE_Main,
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    Mul,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    Ent1,
    HEnt1,
    HMsgBox,
    UTOF,
    HTB97,
    Buttons,
    Hcompte,
    SAISUTIL,
    Menus,
    AGLInit,       // TheData
    CRITEDT,
    {$IFDEF MODENT1}
    CPTypeCons,
    {$ENDIF MODENT1}
    UTOB
    ;

//==================================================
// Externe
//==================================================
procedure ZoomEcritureMvt(lequel : String ; quoi : TFichierBase ; LaListe : String);

//==================================================
// Definition de class
//==================================================
Type
    TOF_ZECRIMVT = Class (TOF)
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
        // parametre
        lequel : string;
        quoi : TFichierBase;

        // control Sup
        PopZ : TPopupMenu;

        // compte general ??
        CGen : TGGeneral;

        // edition
        ACritEdt : ClassCritEdt;

        procedure OnDblClickFListe(Sender: TObject);
        procedure OnChangeE_EXERCICE(Sender: TObject);

        procedure OnEnterBMenuZoom(Sender: TObject);
        procedure OnClickBBJustif(Sender: TObject);
        procedure OnClickBEche(Sender: TObject);
{$IFDEF EDTQR}
        procedure OnClickBGLAge(Sender: TObject);
        procedure OnClickBGlVenPre(Sender: TObject);
        procedure OnClickBGlVenRet(Sender: TObject);
        procedure OnClickBJalCpt(Sender: TObject);
{$ENDIF}
        procedure OnClickBGLGenANA(Sender: TObject);
        procedure OnClickBJalDiv(Sender: TObject);
        procedure OnClickBGL(Sender: TObject);
        procedure OnClickBVisuP(Sender: TObject);

        procedure InitCriteres;
        procedure FormAfterShow;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

Uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
    UtilEdt,
{$IFNDEF IMP}
    UTOFCPGLGENE,
    UTOFCPGLAUXI,
    TofEcheancier, {JP 13/10/05 : FQ 16861}
    UTOFCPJALECR, {JP 14/10/05 : FQ 16864}
    SaisBor,
{$ENDIF}
{$IFNDEF EAGLCLIENT}
  {$IFNDEF IMP}
    {$IFNDEF CCS3}
{$IFDEF EDTQR}
    QRGLAge,
    QRGLVen,
      {$ENDIF}
    {$ENDIF}
{$IFDEF EDTQR}
    QRJCpte,
    QRGLGESE,
    {$ENDIF}
  {$ENDIF}
{$ELSE}
{$ENDIF}
    Saisie,
    FILTRE,
    UtilPGI
    ;

//==================================================
// Definition des Variables
//==================================================
var
    Mess : array[0..2] of string  = (
    {00}    'Détail des mouvements du compte général :',
    {01}    'Détail des mouvements du compte tiers :',
    {02}    'Détail des mouvements du journal :'
            );

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/02/2003
Modifié le ... : 24/02/2003
Description .. : Recupere une chaine a partire d'un mode d'action
Mots clefs ... :
*****************************************************************}
function TFToStr(TF : TFichierBase) : String;
begin
    if (TF = fbGene)      then result := 'GENERAUX'
    else if (TF = fbJal)  then result := 'JOURNAL'
    else if (TF = fbAux)  then result := 'TIERS'
    else result := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/02/2003
Modifié le ... : 24/02/2003
Description .. : Recupere une action a partire d'une chaine
Mots clefs ... :
*****************************************************************}
function StrToTF(psz : String) : TFichierBase;
begin
    if (psz = 'GENERAUX')     then result := fbGene
    else if (psz = 'JOURNAL') then result := fbJal
    else if (psz = 'TIERS')   then result := fbAux
    else result := fbNone;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure ZoomEcritureMvt(lequel : String ; quoi : TFichierBase ; LaListe : String);
begin
    if _Blocage(['nrCloture'],true,'nrAucun') then exit;
    if (not (quoi in [fbGene,fbJal,fbAux])) then exit;

    AGLLanceFiche('CP','ZECRIMVT','','','ACTION=CONSULTATION;' + lequel + ';' + LaListe + ';' + TFToStr(quoi));
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnNew;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnDelete;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnUpdate;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnLoad;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnArgument(S : String);
var
//    D1,D2 : TDateTime;
    S1,S2 : string;
    liste : string;
    FListe : THGrid;
    E_EXERCICE : THValComboBox;
    BMenuZoom : TToolBarButton97;
    bouton : TBitBtn;
begin
    Inherited;

    ACritEdt := ClassCritEdt.Create;

    // init des cntrol sup
    POPZ := TPopupMenu.Create(Ecran);
    CGen := nil;

//    MemoStyle := msBook;

    // recuperation des parametre
    S1 := UpperCase(S);
    S2 := ReadTokenSt(S1);                      // mode eAGL
    S2 := ReadTokenSt(S1); lequel := S2;        // le compte
    S2 := ReadTokenSt(S1); liste := S2;         // la liste
    S2 := ReadTokenSt(S1); quoi := StrToTF(S2); // quoi

    // recuperationn des item et assign des fct
    FListe := THGrid(GetControl('FListe',true)); if (assigned(FListe)) then FListe.OnDblClick := OnDblClickFListe;
    E_EXERCICE := THValComboBox(GetControl('E_EXERCICE',true)); if (assigned(E_EXERCICE)) then E_EXERCICE.OnChange := OnChangeE_EXERCICE;
    // zoom
    BMenuZoom := TToolBarButton97(GetControl('BMenuZoom',true)); if (assigned(BMenuZoom)) then begin BMenuZoom.OnMouseEnter := OnEnterBMenuZoom; BMenuZoom.DropdownMenu := POPZ; end;
    bouton := TBitBtn(GetControl('BBJustif',true)); if (assigned(bouton)) then bouton.OnClick := OnClickBBJustif;
    bouton := TBitBtn(GetControl('BEche',true)); if (assigned(bouton)) then bouton.OnClick := OnClickBEche;

    {$IFDEF EDTQR}
    bouton := TBitBtn(GetControl('BGLAge',true));
    if (assigned(bouton)) then bouton.OnClick := OnClickBGLAge;
    {$ENDIF}
    {JP 13/10/05 : FQ 16861 : à ma connaissance, le Grand livre Agé n'est pas passé en non QR}
    Bouton.Tag := 0;

    bouton := TBitBtn(GetControl('BGLGenANA',true)); if (assigned(bouton)) then bouton.OnClick := OnClickBGLGenANA;

    {$IFDEF EDTQR}
    bouton := TBitBtn(GetControl('BGlVenPre',true));
    if (assigned(bouton)) then bouton.OnClick := OnClickBGlVenPre;
    {$ENDIF}
    {$IFDEF EAGL}
    {JP 13/10/05 : FQ 16861 : à ma connaissance, le Grand livre ventilé des prévisions n'est pas passé en eAgl}
    Bouton.Tag := 0;
    {$ENDIF EAGL}

    bouton := TBitBtn(GetControl('BGlVenRet',true));
    {$IFDEF EDTQR}
    if (assigned(bouton)) then bouton.OnClick := OnClickBGlVenRet;
    {$ENDIF}
    {JP 13/10/05 : FQ 16861 : à ma connaissance, le Grand livre ventilé des retards n'est pas passé en non QR}
    Bouton.Tag := 0;

    bouton := TBitBtn(GetControl('BJalCpt',true));
    {$IFDEF EDTQR}
    if (assigned(bouton)) then bouton.OnClick := OnClickBJalCpt;
    {$ENDIF}
    {JP 14/10/05 : FQ 16864 : à ma connaissance, le Journal par compte général n'est pas passé en non QR}
    Bouton.Tag := 0;

    bouton := TBitBtn(GetControl('BJalDiv',true)); if (assigned(bouton)) then bouton.OnClick := OnClickBJalDiv;
    bouton := TBitBtn(GetControl('BGL',true)); if (assigned(bouton)) then bouton.OnClick := OnClickBGL;
    bouton := TBitBtn(GetControl('BVisuP',true)); if (assigned(bouton)) then bouton.OnClick := OnClickBVisuP;

    // initialisation
    InitCriteres;
    LibellesTableLibre(TTabSheet(GetControl('ZONESLIBRES',true)),'TE_TABLE','E_TABLE','E');
    SetControlProperty('E_TABLE0','DataType','TZNATECR0');
    SetControlProperty('E_TABLE1','DataType','TZNATECR1');
    SetControlProperty('E_TABLE2','DataType','TZNATECR2');
    SetControlProperty('E_TABLE3','DataType','TZNATECR3');

    // enable des btn zoom
    SetControlEnabled('BGLGenANA',false);
    SetControlEnabled('BBJustif',false);
    SetControlEnabled('BEche',false);
    SetControlEnabled('BGLAge',false);
    SetControlEnabled('BGlVenRet',false);
    SetControlEnabled('BGlVenPre',false);
    SetControlEnabled('BJalDiv',false);
    SetControlEnabled('BJalCpt',false);

    Case quoi of
        fbGene :
            begin
                SetControlVisible('TE_GENERAL',false);
                SetControlVisible('E_GENERAL',false);
                SetControlText('E_GENERAL',lequel);

                SetControlProperty('TE_AUXILIAIRE','Left',GetControl('TE_GENERAL').Left);
                SetControlProperty('TE_AUXILIAIRE','Top',GetControl('TE_GENERAL').Top);
                SetControlProperty('E_AUXILIAIRE','Left',GetControl('E_GENERAL').Left);
                SetControlProperty('E_AUXILIAIRE','TabOrder',5);
                SetControlProperty('E_AUXILIAIRE','Top',GetControl('E_GENERAL').Top);

                Ecran.Caption := Mess[0] + lequel ;
                Ecran.HelpContext := 7109070 ;

                if lequel <> '' then
                begin
                    CGen := TGGeneral.Create(lequel);
                    if (Cgen <> nil) then if (CGen.Ventilable[1] or CGen.Ventilable[2] or CGen.Ventilable[3] or CGen.Ventilable[4] or CGen.Ventilable[5]) then SetControlEnabled('BGLGenANA',true);
                end;
            end;
        fbAux :
            begin
                SetControlVisible('TE_AUXILIAIRE',false);
                SetControlVisible('E_AUXILIAIRE',false);
                SetControlText('E_AUXILIAIRE',lequel);

                Ecran.Caption := Mess[1] + lequel;
                Ecran.HelpContext := 7142070 ;

                SetControlEnabled('BBJustif',true);
                SetControlEnabled('BEche',true);
                SetControlEnabled('BGLAge',true);
                SetControlEnabled('BGlVenRet',true);
                SetControlEnabled('BGlVenPre',true);
            end;
        fbJal :
            begin
                SetControlProperty('E_JOURNAL','Value',lequel);
                Ecran.Caption := Mess[2] + lequel;
                Ecran.HelpContext:=7208060 ;

                SetControlEnabled('BGL',false);
                SetControlEnabled('BJalDiv',true);
                SetControlEnabled('BJalCpt',true);

                if (GetControlText('E_JOURNAL') = 'CLO') then
                begin
                    SetControlEnabled('BJalDiv',false);
                    SetControlEnabled('BJalCpt',false);
                end;
            end;
    end;

    UpdateCaption(Ecran);

    {JP 02/11/05 : FQ 16962 : Gérer dans InitCriteres
    SetControlText('E_EXERCICE',QuelDateExo(V_PGI.DateEntree,D1,D2));
    ExoToDates(GetControlText('E_EXERCICE'),GetControl('E_DATECOMPTABLE',true),GetControl('E_DATECOMPTABLE_',true));}
    SetControlText('E_DEVISE',V_PGI.DevisePivot);

    PositionneEtabUser(THValComboBox(GetControl('E_ETABLISSEMENT',true)));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnClose;
begin
    FreeAndNil(POPZ);
    FreeAndNil(CGen);
    FreeAndNil(ACritEdt);

    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnDisplay;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnCancel;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnChangeE_EXERCICE(Sender : TObject);
begin
    {JP 14/08/07 : FQ 20091 : Nouvelle fonction dans Ent1, qui gère le cas où Exo = ''}
    ExoToEdDates(GetControlText('E_EXERCICE'),GetControl('E_DATECOMPTABLE',true),GetControl('E_DATECOMPTABLE_',true));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnDblClickFListe(Sender : TObject);
{$IFDEF EAGLCLIENT}
var
    TOBListe,TOBLigne : TOB;
    Index             : Integer;     {FP 04/11/2005 FQ16858}
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
    TOBListe := TOB.Create('Liste Ecriture',nil,-1);
    {b FP 04/11/2005 FQ16858}
    Index := THGrid(GetControl('FListe',true)).row-1;
    if (Index >= 0) and (Index < GetDataSet.Detail.Count) then
    begin
      TOBLigne := GetDataSet.Detail[Index];
      TOB.Create('ECRITURE',TOBListe,-1);
      TOBListe.Detail[0].Dupliquer(TOBLigne,False,True);
    end
    else
    begin
      TOB.Create('ECRITURE',TOBListe,-1);
    end;
    {e FP 04/11/2005 FQ16858}
  {$IFNDEF IMP}
    if (PieceSurFolio(TOBListe.FindField('E_JOURNAL').AsString)) then {LanceSaisieFolio(TOBListe,taConsult)}
    else
  {$ENDIF}
    TrouveEtLanceSaisie(TOBListe,taConsult,'');
    FreeAndNil(TOBListe);
{$ELSE}
  {$IFNDEF IMP}
    if (PieceSurFolio(TFMul(Ecran).Q.FindField('E_JOURNAL').AsString)) then LanceSaisieFolio(TFMul(Ecran).Q,taConsult)
    else
  {$ENDIF}
    TrouveEtLanceSaisie(TFMul(Ecran).Q,taConsult,'');
{$ENDIF}
end;

{ Procédures de zoom sur les états à partir de la fiche }

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnEnterBMenuZoom(Sender : TObject);
begin
    PopZoom97(TToolBarButton97(GetControl('BMenuZoom',true)),POPZ);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnClickBGLGenANA(Sender : TObject);
{$IFDEF EAGLCLIENT}
{$ELSE}
Var
    Crit : TCritEdt;
    D1,D2 : TdateTime ;
    Axe : String3 ;
    Machin : String ;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    if (CGen = nil) then exit;

    Axe := 'A1';

    if (CGen.Ventilable[1])      then Axe := 'A1'
    else if (CGen.Ventilable[2]) then Axe := 'A2'
    else if (CGen.Ventilable[3]) then Axe := 'A3'
    else if (CGen.Ventilable[4]) then Axe := 'A4'
    else if (CGen.Ventilable[5]) then Axe := 'A5'
    else exit;

    Fillchar(Crit,SizeOf(Crit),#0);
    D1 := StrToDate(GetControlText('E_DATECOMPTABLE'));
    D2 := StrToDate(GetControlText('E_DATECOMPTABLE_'));
    Crit.Date1 := D1;
    Crit.Date2 := D2;
    Crit.DateDeb := Crit.Date1;
    Crit.DateFin := Crit.Date2;
    Crit.NatureEtat := neGL;
    Crit.GL.Axe := Axe;

    InitCritEdt(Crit);

    Crit.Cpt1 := CGen.General;
    Crit.Cpt2 := Crit.Cpt1;
    Machin := TFMul(Ecran).Q.CRITERES;
    while (Pos('E_', Machin) > 0) do Machin[Pos('E_',Machin)] := 'Y';
    Crit.SQLPLUS := 'AND ' + Machin;

{$IFNDEF IMP}
{$IFDEF EDTQR}
    GLGESEZoom(Crit);
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... : 13/10/2005
Description .. : JP : 16861 : branchement du justificatif de solde en eAgl
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnClickBBJustif(Sender : TObject);
begin
  {JP 13/10/05 : FQ 16861 : Branchement de du justificatif de solde}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  ACritEdt.CritEdt.Date1 := StrToDate(GetControlText('E_DATECOMPTABLE'));
  ACritEdt.CritEdt.Date2 := StrToDate(GetControlText('E_DATECOMPTABLE_'));
  ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
  ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;
  ACritEdt.CritEdt.NatureEtat := neGL;
  ACritEdt.CritEdt.GL.EnDateSituation := True;
  ACritEdt.CritEdt.Cpt1 := GetControlText('E_AUXILIAIRE');
  ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;
  ACritEdt.CritEdt.Etab := GetControlText('E_ETABLISSEMENT');
  ACritEdt.CritEdt.DeviseSelect := GetControlText('E_DEVISE');
  ACritEdt.CritEdt.Ech.Jal1 := GetControlText('E_JOURNAL');
  ACritEdt.CritEdt.Ech.Jal2 := GetControlText('E_JOURNAL_');
  ACritEdt.CritEdt.Qualifpiece := GetControlText('E_QUALIFPIECE');
  {JP 19/10/05 : FQ 16911 : Petite protection qui manquait}
  if GetControlText('E_NUMEROPIECE' ) = '' then ACritEdt.CritEdt.GL.NumPiece1 := 0
                                           else ACritEdt.CritEdt.GL.NumPiece1 := StrToInt(GetControlText('E_NUMEROPIECE'));
  if GetControlText('E_NUMEROPIECE_') = '' then ACritEdt.CritEdt.GL.NumPiece2 := 999999999
                                           else ACritEdt.CritEdt.GL.NumPiece2 := StrToInt(GetControlText('E_NUMEROPIECE_'));

  TheData := ACritEdt;

  {$IFNDEF IMP}
  CPLanceFiche_CPGLAUXI('');
  {$ENDIF}
  TheData := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnClickBEche(Sender : TObject);
begin
  {JP 13/10/05 : FQ 16861 : Branchement de l'échéancier}
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  ACritEdt.CritEdt.Date1 := StrToDate(GetControlText('E_DATEECHEANCE'));
  ACritEdt.CritEdt.Date2 := StrToDate(GetControlText('E_DATEECHEANCE_'));
  ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
  ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;
  ACritEdt.CritEdt.Ech.Cpt3 := GetControlText('E_GENERAL');
  ACritEdt.CritEdt.Ech.Cpt4 := ACritEdt.CritEdt.Ech.Cpt3;
  ACritEdt.CritEdt.NatureEtat := neEch;
  ACritEdt.CritEdt.Cpt1 := GetControlText('E_AUXILIAIRE');
  ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;

  ACritEdt.CritEdt.Etab := GetControlText('E_ETABLISSEMENT');
  ACritEdt.CritEdt.DeviseSelect := GetControlText('E_DEVISE');
  ACritEdt.CritEdt.Ech.Jal1 := GetControlText('E_JOURNAL');
  ACritEdt.CritEdt.Ech.Jal2 := GetControlText('E_JOURNAL_');
       if GetControlText('E_QUALIFPIECE') = 'N' then ACritEdt.CritEdt.Qualifpiece := 'N'
  else if GetControlText('E_QUALIFPIECE') = 'S' then ACritEdt.CritEdt.Qualifpiece := 'S'
  else if GetControlText('E_QUALIFPIECE') = 'U' then ACritEdt.CritEdt.Qualifpiece := 'U'
  else if GetControlText('E_QUALIFPIECE') = ''  then ACritEdt.CritEdt.Qualifpiece := 'NSU';

  TheData := ACritEdt;

  {$IFNDEF IMP}
  AGLLanceFiche('CP', 'EPECHEANCIER','','ECH', 'ECH') ;
  {$ENDIF}
  TheData := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EDTQR}
procedure TOF_ZECRIMVT.OnClickBGLAge(Sender : TObject);
{$IFDEF EAGLCLIENT}
{$ELSE}
Var
    Crit : TCritEdt;
    D1 : TdateTime;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    Fillchar(Crit,SizeOf(Crit),#0);
    D1 := StrToDate(GetControlText('E_DATECOMPTABLE_'));
    Crit.Date1 := D1;
    Crit.DateDeb := Crit.Date1;
    Crit.NatureEtat := neGlV;

    InitCritEdt(Crit);

    Crit.Cpt1 := GetControlText('E_AUXILIAIRE');
    Crit.Cpt2 := Crit.Cpt1;
    Crit.SCpt1 := GetControlText('E_GENERAL');
    Crit.SCpt2 := Crit.SCpt1;
    Crit.GlV.TriePar := 0;
    Crit.GlV.TypePar := 0;
    Crit.SautPage := 2;
    Crit.GlV.ChoixEcart := 0;
    Crit.SQLPLUS := ' AND ' + TFMul(Ecran).Q.CRITERES + ' ';

{$IFNDEF IMP}
{$IFNDEF CCS3}
    GLivreAgeZoom(Crit);
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;
{$ENDIF}
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EDTQR}
procedure TOF_ZECRIMVT.OnClickBGlVenRet(Sender : TObject);
{$IFDEF EAGLCLIENT}
{$ELSE}
Var
    Crit : TCritEdt;
    D1 : TdateTime;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    Fillchar(Crit,SizeOf(Crit),#0);
    D1 := StrToDate(GetControlText('E_DATECOMPTABLE_'));
    Crit.Date1 := D1;
    Crit.DateDeb := Crit.Date1;
    Crit.NatureEtat := neGlV;

    InitCritEdt(Crit);

    Crit.Cpt1 := GetControlText('E_AUXILIAIRE');
    Crit.Cpt2 := Crit.Cpt1;
    Crit.SCpt1 := GetControlText('E_GENERAL');
    Crit.SCpt2 := Crit.SCpt1;
    Crit.GlV.TriePar := 0;
    Crit.GlV.TypePar := 0;
    Crit.SautPage := 2;
    Crit.Glv.TypeGL := 1;
    Crit.GlV.ChoixEcart := 0;
    Crit.SQLPLUS := ' AND ' + TFMul(Ecran).Q.CRITERES + ' ';

{$IFNDEF IMP}
{$IFNDEF CCS3}
    GLVentileZoom (Crit) ;
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;
{$ENDIF}
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EDTQR}
procedure TOF_ZECRIMVT.OnClickBGlVenPre(Sender : TObject);
{$IFDEF EAGLCLIENT}
{$ELSE}
Var
    Crit : TCritEdt;
    D1 : TdateTime;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    Fillchar(Crit,SizeOf(Crit),#0);
    D1 := StrToDate(GetControlText('E_DATECOMPTABLE_'));
    Crit.Date1 := D1;
    Crit.DateDeb := Crit.Date1;
    Crit.NatureEtat := neGlV;

    InitCritEdt(Crit);

    Crit.Cpt1 := GetControlText('E_AUXILIAIRE');
    Crit.Cpt2 := Crit.Cpt1;
    Crit.SCpt1 := GetControlTExt('E_GENERAL');
    Crit.SCpt2 := Crit.SCpt1;
    Crit.GlV.TriePar := 0;
    Crit.GlV.TypePar := 0;
    Crit.SautPage := 2;
    Crit.Glv.TypeGL := 0;
    Crit.GlV.ChoixEcart := 0;
    Crit.SQLPLUS := ' AND ' + TFMul(Ecran).Q.CRITERES + ' ';

{$IFNDEF IMP}
{$IFNDEF CCS3}
    GLVentileZoom(Crit);
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnClickBJalDiv(Sender : TObject);
begin
  {JP 14/10/05 : FQ 16864 : branchement du Journal des écritures (Divisionnaire) en eAgl}
  Fillchar(ACritEdt.CritEdt,SizeOf(ACritEdt.CritEdt),#0);
  ACritEdt.CritEdt.Exo.Code := GetControlText('E_EXERCICE');
  ACritEdt.CritEdt.Date1 := StrToDate(GetControlText('E_DATECOMPTABLE'));
  ACritEdt.CritEdt.Date2 := StrToDate(GetControlText('E_DATECOMPTABLE_'));
  ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
  ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;
  ACritEdt.CritEdt.NatureEtat := neJAl;
  ACritEdt.CritEdt.Cpt1 := GetControlText('E_JOURNAL');
  ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;
  if GetControlText('E_QUALIFPIECE') <> '' then
    ACritEdt.CritEdt.Qualifpiece := GetControlText('E_QUALIFPIECE') + ';';
  ACritEdt.CritEdt.Etab := GetControlText('E_ETABLISSEMENT');
  if GetControlText('E_NUMEROPIECE' ) <> '' then ACritEdt.CritEdt.Jal.NumPiece1 := StrToInt(GetControlText('E_NUMEROPIECE'))
                                            else ACritEdt.CritEdt.Jal.NumPiece1 := 0;
  if GetControlText('E_NUMEROPIECE_') <> '' then ACritEdt.CritEdt.Jal.NumPiece2 := StrToInt(GetControlText('E_NUMEROPIECE_'))
                                            else ACritEdt.CritEdt.Jal.NumPiece2 := 999999999;
  ACritEdt.CritEdt.DeviseSelect := GetControlText('E_DEVISE');
  TheData := ACritEdt;
  {$IFNDEF IMP}
  CPLanceFiche_CPJALECR;
  {$ENDIF}
  TheData := nil; {Par précaution, car théoriquement, c'est déjà fait}
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFDEF EDTQR}
procedure TOF_ZECRIMVT.OnClickBJalCpt(Sender : TObject);
{$IFDEF EAGLCLIENT}
{$ELSE}
Var
    Crit : TCritEdt;
    D1,D2 : TdateTime;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
{$ELSE}
    Fillchar(Crit,SizeOf(Crit),#0);
    D1 := StrToDate(GetControlText('E_DATECOMPTABLE'));
    D2 := StrToDate(GetControlText('E_DATECOMPTABLE_'));
    Crit.Date1 := D1;
    Crit.Date2 := D2;
    Crit.DateDeb := Crit.Date1;
    Crit.DateFin := Crit.Date2;
    Crit.NatureEtat := neJAl;

    InitCritEdt(Crit);

    Crit.Cpt1 := GetControlText('E_JOURNAL');
    Crit.Cpt2 := Crit.Cpt1;

{$IFNDEF IMP}
{$IFDEF EDTQR}
    JalCpteGeZoom(Crit);
{$ENDIF}
{$ENDIF}
{$ENDIF}
end;
{$ENDIF}
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnClickBGL(Sender : TObject);
Var
    D1,D2 : TdateTime;
begin
    Fillchar(ACritEdt.CritEdt,SizeOf(ACritEdt.CritEdt),#0);
    D1 := StrToDate(GetControlText('E_DATECOMPTABLE'));
    D2 := StrToDate(GetControlText('E_DATECOMPTABLE_'));
    ACritEdt.CritEdt.Date1 := D1;
    ACritEdt.CritEdt.Date2 := D2;
    ACritEdt.CritEdt.DateDeb := ACritEdt.CritEdt.Date1;
    ACritEdt.CritEdt.DateFin := ACritEdt.CritEdt.Date2;
    ACritEdt.CritEdt.NatureEtat := neGL;

    case quoi of
        fbGene :
            begin
                ACritEdt.CritEdt.GL.ForceNonCentralisable := true;
                ACritEdt.CritEdt.Cpt1 := GetControlText('E_GENERAL');
                ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;
                ACritEdt.CritEdt.SQLPLUS := ' AND ' + TFMul(Ecran).Q.CRITERES + ' ';

{$IFNDEF IMP}
                TheData := ACritEdt;
                CPLanceFiche_CPGLGene('');
                TheData := nil;
{$ENDIF}
            end;
        fbAux :
            begin
                ACritEdt.CritEdt.Cpt1 := GetControlText('E_AUXILIAIRE');
                ACritEdt.CritEdt.Cpt2 := ACritEdt.CritEdt.Cpt1;
                ACritEdt.CritEdt.SQLPLUS := 'AND ' + TFMul(Ecran).Q.CRITERES + ' ';

{$IFNDEF IMP}
                TheData := ACritEdt;
                CPLanceFiche_CPGLAuxi('');
                TheData := nil;
{$ENDIF}
            end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.OnClickBVisuP(Sender : TObject);
begin
    OnDblClickFListe(Sender);
end;

{ Fin des Zooms }

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ZECRIMVT.InitCriteres;
begin
    {JP 02/11/05 : FQ 16962 : Affectation de l'exercice et des dates à partir de la date d'entrée
    SetControlText('E_EXERCICE', QuelExoDtBud(V_PGI.DateEntree));
    ExoToDates(GetControlText('E_EXERCICE'),GetControl('E_DATECOMPTABLE', True), GetControl('E_DATECOMPTABLE_', True));
    
    JP 14/08/07 : FQ 20091 : le code ci-dessus est déplacé dans le OnAfterFormShow, car l'agl intervient sur les dates}
    TFMul(Ecran).OnAfterFormShow := FormAfterShow;

    SetControlText('E_DATEECHEANCE',StDate1900);
    SetControlText('E_DATEECHEANCE_',StDate2099);
    SetControlText('E_DATECREATION',StDate1900);
    SetControlText('E_DATECREATION_',StDate2099);

    Case quoi of
        fbGene : SetControlText('E_GENERAL',Lequel);
        fbAux  : SetControlText('E_AUXILIAIRE',Lequel);
        fbJal  : SetControlText('E_JOURNAL',Lequel);
    end;
end;

{JP 14/08/07 : FQ 20091 : L'agl, au moins en CWAS, initialise les dates avec leur valeur par défaut
               après le OnArgument.
{---------------------------------------------------------------------------------------}
procedure TOF_ZECRIMVT.FormAfterShow;
{---------------------------------------------------------------------------------------}
begin
  {JP 02/11/05 : FQ 16962 : Affectation de l'exercice et des dates à partir de la date d'entrée}
  SetControlText('E_EXERCICE', QuelExoDtBud(V_PGI.DateEntree));
  ExoToEdDates(GetControlText('E_EXERCICE'),GetControl('E_DATECOMPTABLE', True), GetControl('E_DATECOMPTABLE_', True));
end;

//================================================================================
// Initialization
//================================================================================

Initialization
    registerclasses([TOF_ZECRIMVT]);
end.
