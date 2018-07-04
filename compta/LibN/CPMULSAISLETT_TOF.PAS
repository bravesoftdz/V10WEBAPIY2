{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... : 17/03/2003
Description .. : Source TOF de la FICHE : CPMULSAISLETT ()
Mots clefs ... : TOF;CPMULSAISLETT
*****************************************************************}
Unit CPMULSAISLETT_TOF;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls,
    Controls,
    ExtCtrls,
    Classes,
{$IFDEF EAGLCLIENT}
    MaineAGL,
    eMUL,
{$ELSE}
    db,
    hdb,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    FE_Main,
    MUL,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    Ent1,
    HEnt1,
    HMsgBox,
    UTOF,
    SaisUtil,
    LettUtil,
    utilPGI
    ;

//==================================================
// Externe
//==================================================
procedure LettrerEnSaisie(X : RMVT ; NbEche : integer = 0 ; StWhere : String = '');
procedure LettrerEnSaisieTreso(X : RMVT ; Gen,Aux : String ; SansLesPartiel : Boolean);
procedure DelettrerEnSaisieTreso(X : RMVT ; Gen,Aux,CodeLettre : String ; SansLesPartiel : Boolean);

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPMULSAISLETT = Class (TOF)
    public
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
        procedure OnTimer(Sender: TObject);
        procedure OnDblClickFListe(Sender: TObject);
    public
    private
        TT : TTimer;
        NbEche : integer;
        { FQ 19816 BVE 02.05.07 }
        gLoad : boolean;
        gDisplay : boolean;
        gUpdate : boolean;
        { END FQ 19816 }
{$IFDEF EAGLCLIENT}
        FListe : THGrid;
{$ELSE EAGLCLIENT}
        FListe : THDBGrid;
{$ENDIF EAGLCLIENT}
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
    {$IFDEF MODENT1}
    CPProcGen,
    {$ENDIF MODENT1}
    Lettrage
    ,Variants
    ;

//==================================================
// Definition des Variable
//==================================================
var
    MVT : RMVT;
    HLS : array [0..2] of string = (
        '0;?caption?;Lettrage non autorisé : cette échéance est déjà totalement lettrée.;W;O;O;O;',
        'Distinguer',
        '2'
        );

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LettrerEnSaisie(X : RMVT ; NbEche : integer ; StWhere : String);
begin
    MVT := X;
    AGLLanceFiche('CP','CPMULSAISLETT','','','ACTION=CONSULTATION;' + IntToStr(NbEche) + ';' + StWhere);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LettrerEnSaisieTreso( X : RMVT ; Gen,Aux : String ; SansLesPartiel : Boolean) ;
var
    R : RLETTR;
begin
    if (not X.Treso) then exit;

    FillChar(R,Sizeof(R),#0) ;
    R.General := Gen;
    R.Auxiliaire := Aux;
    R.Appel := tlSaisieTreso;
    R.GL := Nil;
    R.CritMvt := ' AND E_GENERAL="'+R.General+'" AND E_AUXILIAIRE="'+R.Auxiliaire+'" AND E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI" ';
    R.CritDev := X.CodeD;
    R.LettrageDevise := (X.CodeD <> V_PGI.DevisePivot);
    R.DeviseMvt := X.CodeD;
    R.Ident := X;
    R.Ident.NumLigne := X.NumLigne;
    R.Ident.NumEche := X.NumEche;
    R.SansLesPartiel := SansLesPartiel;

    LettrageManuel(R,true,taModif);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure DelettrerEnSaisieTreso( X : RMVT ; Gen,Aux,CodeLettre : String ; SansLesPartiel : Boolean) ;
var
    R : RLETTR;
begin
    if (not X.Treso) then exit;

    FillChar(R,Sizeof(R),#0);
    R.General := Gen;
    R.Auxiliaire := Aux;
    R.Appel := tlSaisieTreso;
    R.GL := Nil;
    R.CritMvt := ' AND E_GENERAL="'+R.General+'" AND E_AUXILIAIRE="'+R.Auxiliaire+'" AND E_ETATLETTRAGE<>"RI" ';
    R.CritDev := X.CodeD;
    R.LettrageDevise := (X.CodeD <> V_PGI.DevisePivot);
    R.DeviseMvt := X.CodeD;
    R.Ident := X;
    R.Ident.NumLigne := X.NumLigne;
    R.Ident.NumEche := X.NumEche;
    R.SansLesPartiel := SansLesPartiel;
    R.CodeLettre := CodeLettre;

    LettrageManuel(R,false,taModif);
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnNew;
begin
    Inherited;
    lastError := 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnDelete;
begin
    Inherited;
    lastError := 1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnUpdate;
begin
    Inherited;
     lastError := 1;
     gUpdate := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnLoad;
begin
    Inherited;
    gLoad := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... : 21/09/2006
Description .. :
Suite ........ : SBO 21/09/2006 : FQ18843 Non compatible 2 tiers...
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnArgument(S : String);
var
    where : string;
    S1,S2 : string;
begin
    Inherited;

    gLoad := false;
    gDisplay := false;
    gUpdate := false;

    // recup des parametre
    S1 := UpperCase(S);
    S2 := ReadTokenSt(S1);                          // mode
    S2 := ReadTokenSt(S1); NbEche := StrToInt(S2);  // nbEcehe
    S2 := ReadTokenSt(S1); where := S2;             // Codition sup

    // set des valeur de mask
    SetControlText('E_JOURNAL',MVT.Jal);
    SetControlText('E_DEVISE',MVT.CodeD);
    SetControlText('E_NUMEROPIECE',IntToStr(MVT.Num));
    SetControlText('E_ETABLISSEMENT',MVT.Etabl);
    SetControlText('E_QUALIFPIECE',MVT.Simul);
    SetControlText('E_EXERCICE',MVT.Exo);

    if (MVT.ModeSaisieJal='BOR') or (MVT.ModeSaisieJal='LIB') then
      begin
      SetControlText('E_DATECOMPTABLE',  StDate1900 );
      SetControlText('E_NATUREPIECE',    '');
      if where <> '' then
        where := where + ' AND ' ;
      where := where + ' E_PERIODE =' + IntToStr( GetPeriode( MVT.DateC ) ) ;
      end
    else
      begin
      SetControlText('E_DATECOMPTABLE',DateToStr(MVT.DateC));
      SetControlText('E_NATUREPIECE',MVT.Nature);
      end ;

    // et du where ...
    SetControlText('XX_WHERE',where);

    // recup des control util
{$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FListe',true));
{$ELSE}
    FListe := THDBGrid(GetControl('FListe',true));
{$ENDIF EAGLCLIENT}
    if (assigned(FListe)) then
      FListe.OnDblClick := OnDblClickFListe;

    // test et
    if (VH^.TenueEuro) then SetControlText(HLS[1],HLS[1] + ' ' + RechDom('TTDEVISETOUTES',V_PGI.DevisePivot,false)+' + '+RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,false));
    if (MVT.CodeD <> V_PGI.DevisePivot) then SetControlVisible('DISTINGUER',false);

  // rafraichissement de la liste
//  TFMul(Ecran).BChercheClick(nil);
    if (NbEche = 1) then
    begin
        TT := TTimer.Create(Ecran);
        TT.OnTimer := OnTimer;
        TT.Enabled := true;
    end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnClose;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnDisplay;
begin
    Inherited;
    gDisplay := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnCancel;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnTimer(Sender: TObject);
begin
    { FQ 19816 BVE 02.05.07 }
    // Le chargement du mul n'est pas fini.
    if not(gLoad and gUpdate and gDisplay) then Exit;
    { END FQ 19816 }
    // desactivation
    TT.Enabled := false;
    TT.Interval := 0;
    // lancement du lettrage
    OnDblClickFListe(Nil);
    // close !
    Ecran.close;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/03/2003
Modifié le ... : 21/09/2006
Description .. :
Suite ........ : SBO 21/09/2006 : FQ18843 Non compatible 2 tiers...
Mots clefs ... :
*****************************************************************}
procedure TOF_CPMULSAISLETT.OnDblClickFListe(Sender: TObject);
var R     : RLETTR;
    CodeL : String ;
begin

  // rien dans le mul ?
  if GetDataSet.Bof and GetDataSet.Eof then Exit;

  // set a blanc
  FillChar(R,Sizeof(R),#0);

  {$IFDEF EAGLCLIENT}
    TFMul(Ecran).Q.TQ.Seek( FListe.Row - 1 ) ;
  {$ENDIF}

  // recup du lettrage s'il y a
  { FQ 19816 BVE 27.04.07 pb Oracle }
  if not VarIsNull(GetField('E_LETTRAGE')) then
     CodeL := GetField('E_LETTRAGE')  // FListe.Cells[6,FListe.Row];
  else
     CodeL := '';
  { END FQ 19816 }

  //verif
  if ((CodeL <> '') and (CodeL = uppercase(CodeL))) then
    begin
    HShowMessage(HLS[0],Ecran.Caption,'');
        exit;
    end;

    // set des variable pour lettrage manuel
    R.General    := GetField('E_GENERAL') ; // FListe.Cells[0,FListe.Row];
    R.Auxiliaire := GetField('E_AUXILIAIRE') ; // FListe.Cells[1,FListe.Row];
    R.Appel      := tlSaisieCour;
    R.GL         := Nil;
    R.CritMvt    := ' AND E_GENERAL="'+R.General+'" AND E_AUXILIAIRE="'+R.Auxiliaire+'" AND E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI" ';
    R.CritDev    := MVT.CodeD;
    R.DeviseMvt  := MVT.CodeD;
    R.LettrageDevise := (MVT.CodeD <> V_PGI.DevisePivot);
    R.Ident          := MVT;
    R.Ident.NumLigne := GetField('E_NUMLIGNE') ; // StrToInt(FListe.Cells[7,FListe.Row]);
    R.Ident.NumEche  := GetField('E_NUMECHE') ; // StrToInt(FListe.Cells[8,FListe.Row]);

    R.Distinguer := (GetCheckBoxState('DISTINGUER') = cbChecked);

    // devise ?
    if (R.DeviseMvt = V_PGI.DevisePivot) then
        // Paquet en Franc ou Euro
        if (R.Distinguer) then
            R.CritDev := V_PGI.DevisePivot;

    // lettrage
    LettrageManuel(R,true,taModif);

    // rafraichissement de la liste
    TFMul(Ecran).BChercheClick(nil);

end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPMULSAISLETT]);
end.
