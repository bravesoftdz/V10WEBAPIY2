{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 17/06/2003
Modifié le ... : 31/03/2004
Description .. : BPY le 31/03/2004 : fonctionnement en 2/3 ....
Suite ........ : merci l'ESP !!
Suite ........ : 
Suite ........ : Source TOF de la FICHE : CPTVATPF ()
Mots clefs ... : TOF;CPTVATPF
*****************************************************************}

Unit CPTVATPF_TOF;

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
{$ELSE}
    FE_MAIN,
    db,
{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    TVAEtab,
{$ENDIF}
    UtilPgi,
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    Ent1,
    HEnt1,
    HMsgBox,
    UTOF,
    UTOB,
    Hpanel,
    HTB97,
    ParamSoc,
    LookUp
    ;

//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
Procedure ParamTVATPF(Tva : boolean);

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPTVATPF = Class (TOF)
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
        TVA : boolean;
        FCodeTVA : string;
        FDossier : string ;

        OnCreat,OnUndo,OnRead,OnWillCreat : boolean;

        HeaderTOB : TOB;
        DetailTOB : TOB;
        TOBGene : TOB;

        FListeMaster : THGrid;
        FListeSlave : THGrid;
        PInfo : THPanel;

        LiRegime : HTStrings;

        first,prev,next,last : TToolbarbutton97;
        GereFARFAE : boolean;
        LesColsFListeSave : string;

        procedure OnRowExitFListeMaster(Sender : TObject ; ou : Integer ; var Cancel : Boolean ; Chg : Boolean);
        procedure OnRowExitFListeSlave(Sender : TObject ; ou : Integer ; var Cancel : Boolean ; Chg : Boolean);
        procedure OnDblClickFListeSlave(Sender: TObject);

        procedure OnClickFirst(Sender : TObject);
        procedure OnClickLast(Sender : TObject);
        procedure OnClickNext(Sender : TObject);
        procedure OnClickPrev(Sender : TObject);

        function  IsModifiedMaster(row : integer) : boolean;
        function  SaveDBMaster(row : integer) : boolean;
        function  CanSaveMaster(row : integer) : boolean;
        function  IsModifiedSlave(row : integer) : boolean;
        function  SaveDBSlave(row : integer) : boolean;
        function  CanSaveSlave(row : integer) : boolean;

        function  SiExisteCode : boolean;
        function  SiExisteIndice : boolean;

        procedure OnClickBTVAEtab(Sender : TOBject);
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Grids,
  Vierge;

//==================================================
// Definition des Constant
//==================================================
Const Messages : Array[0..20] of String =(
        {00}    'Codes TVA par régime :',
        {01}    'Codes TPF par régime :',

        'L''enregistrement est inaccessible',

        {03}    'Voulez-vous enregistrer les modifications ?',
        {04}    'Confirmez-vous la suppression de l''enregistrement ?',

        {05}    'Vous devez renseigner un code.',
        {06}    'Vous devez renseigner un libellé.',
        {07}    'Le code que vous avez saisi existe déjà. Vous devez le modifier.',
        {08}    'Cet indice de code TVA est déjà attribué.',

        '7;?caption?;Vous devez renseigner le taux à l''achat.;W;O;O;O;',
        '8;?caption?;Vous devez renseigner le taux à la vente.;W;O;O;O;',
        {11}    'Vous devez renseigner le compte à l''achat.',
        {12}    'Vous devez renseigner le compte à la vente.',

        '11;?caption?;L''enregistrement ne sera pas sauvegardé.;W;O;O;O;',
        {14}    'Le code est référencé par un compte général.',

        {15}    'Le compte à l''achat débit que vous avez saisi n''existe pas, est incompatible ou incomplet.',
        {16}    'Le compte à la vente débit que vous avez saisi n''existe pas, est incompatible ou incomplet.',
        {17}    'Le compte à l''achat encaissement que vous avez saisi n''existe pas, est incompatible ou incomplet.',
        {18}    'Le compte à la vente encaissement que vous avez saisi n''existe pas, est incompatible ou incomplet.',
        {19}    'Le compte de TVA de R.G. que vous avez saisi n''existe pas, est incompatible ou incomplet.',
        {20}    'N''oubliez pas que la TVA et la TPF doivent avoir le même code.'
        );

//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
Procedure ParamTVATPF(Tva : boolean);
begin
    if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrEnca','nrDeca'],True,'nrBatch') then exit;

    if (TVA) then AGLLanceFiche('CP','CPTVATPF','','','ACTION=MODIFICATION;TVA')
    else AGLLanceFiche('CP','CPTVATPF','','','ACTION=MODIFICATION;TPF');

    _Bloqueur('nrBatch',False);
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTVATPF.OnNew;
var
    cancel : boolean;
    tmp : TOB;
begin
    Inherited;

    // en mose creation ??
    if (OnCreat = true) then exit;

    // Positionnement
    SetControlEnabled('CC_CODE',true);
    SetFocusControl('CC_CODE');

    // Ajout d'une ligne dans la Grille
    if (HeaderTOB.Detail.Count >= (FListeMaster.RowCount-1)) then FListeMaster.RowCount := FListeMaster.RowCount+1
    else if (HeaderTOB.Detail.Count = 0) then
    else exit;

    //tob fille sup !
    tmp := TOB.Create('CHOIXCOD',HeaderTOB,-1);
    tmp.InitValeurs;
    tmp.PutValue('CC_LIBRE','0') ;

    // set de la variable de précreation
    OnWillCreat := true;

    // Positionnement + evt
    if (FListeMaster.RowCount-1 = 1) then OnRowExitFListeMaster(nil,FListeMaster.RowCount-1,cancel,false)
    else FListeMaster.gotoRow(FListeMaster.RowCount-1);

    // set de la variable de creation
    OnCreat := true;

    // disable des ctrl
    SetControlEnabled('bInsert',false);
    SetControlEnabled('bDelete',false);

    // affichage des valeur
    if (TVA) then tmp.PutValue('CC_TYPE','TX1')
    else tmp.PutValue('CC_TYPE','TX2');
    tmp.PutEcran(ecran,PInfo);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTVATPF.OnDelete;
var
    ligne : integer;
    ToDel : TOB;
    Q : TQuery;
begin
    Inherited;

    ligne := FListeMaster.row;
    OnCreat := false;
    OnUndo := false;

    // test : on a le droit et veut on deleté ??
    if (not (PGIAsk(TraduireMemoire(Messages[4]),ecran.caption) = mrYes)) then exit;
    if (TVA) then Q := OpenSql('SELECT G_TVA FROM GENERAUX WHERE G_TVA="' + FListeMaster.Cells[1,ligne] + '"',true)
    else Q := OpenSql('SELECT G_TPF FROM GENERAUX WHERE G_TPF="' + FListeMaster.Cells[1,ligne] + '"',true);
    if (not (Q.Eof)) then
    begin
        PGIBox(Messages[14],Ecran.caption);
        Ferme(Q);
        exit;
    end;
    Ferme(Q);

    // reaffichage
    if (ligne-1 > 0) then FListeMaster.GotoRow(ligne-1)
    else if (ligne+1 >= FListeMaster.RowCount) then OnNew
    else FListeMaster.GotoRow(ligne+1);

    // suppression de la liste actuel
    ToDel := TOB.Create('A Effacé',nil,-1);
    HeaderTOB.Detail[ligne-1].ChangeParent(ToDel,-1);

    // effacement de la liste
    ToDel.DeleteDB;
    FreeAndNil(ToDel);

    // effacement des detail
    if (TVA) then ExecuteSQL('DELETE FROM TXCPTTVA WHERE TV_TVAOUTPF="TX1" AND TV_CODETAUX="' + FListeMaster.Cells[1,ligne] + '"')
    else ExecuteSQL('DELETE FROM TXCPTTVA WHERE TV_TVAOUTPF="TX2" AND TV_CODETAUX="' + FListeMaster.Cells[1,ligne] + '"');

    // effacement de la grille
    FListeMaster.DeleteRow(ligne);

    // disable des ctrl
    first.Enabled := (not (FListeMaster.row = 1));
    prev.Enabled := (not (FListeMaster.Row = 1));
    next.Enabled := (not (FListeMaster.row = (FListeMaster.RowCount - 1)));
    last.Enabled := (not (FListeMaster.row = (FListeMaster.RowCount - 1)));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... : 20/09/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPTVATPF.OnUpdate;
begin
    Inherited;

    // pour les entete
    if (IsModifiedMaster(FListeMaster.row-1) or IsModifiedSlave(FListeSlave.Row-1)) then SaveDBMaster(FListeMaster.row-1);

    // YM 20/09/2005 après modif, maj de la variable globale VH^.LaTOBTVA (FQ 16705)
    ChargeTVATPF; // dans Ent1
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTVATPF.OnLoad;
var
    SQLMaster : string;
    bCancel : boolean;
    lQTva  : TQuery ;
begin
    Inherited;
    bCancel := False;
    if (TVA)
      then SQLMaster := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="TX1"'
      else SQLMaster := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="TX2"';

    lQTva := OpenSelect( SQLMaster, FDossier ) ;
    HeaderTOB := TOB.Create('_CHOIXCOD',nil,-1);
    if not lQTva.Eof then
      HeaderTOB.LoadDetailDB( 'CHOIXCOD', '', '', lQTva, False ) ;
    Ferme( lQTva ) ;

    if (HeaderTOB.Detail.Count > 0) then
    begin
        // maj de l'affichage
        HeaderTOB.PutGridDetail(FListeMaster,false,false,'CC_CODE;CC_LIBELLE',true);
        OnRowExitFListeMaster(nil,0,bCancel,true);
    end
    // créé un new :)
    else OnNew;

    OnRead := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTVATPF.OnArgument(S : String);
var
    s1,s2 : string;
    QLoc : TQuery;
    BTVAEtab : TToolBarButton97;
begin
    Inherited;

    // mode ...
    OnWillCreat := false;
    OnCreat := false;
    OnRead := true;
    OnUndo := false;

    if EstTablePartagee('TXCPTTVA')
      then FDossier := TableToBase( 'TXCPTTVA' )
      else FDossier := '' ;

    s1 := UpperCase(S);
    s2 := ReadTokenSt(s1);                      // ACTION ....
    s2 := ReadTokenSt(s1);
    TVA := (s2 = 'TVA'); // TVA ou TPF ....
{$IFNDEF GCGC}
    GereFARFAE := False;
{$ELSE GCGC}
    GereFARFAE := True;
{$ENDIF GCGC}

    // recup de la grille
    FListeMaster := THGrid(GetControl('FLISTEMASTER',true));
    FListeMaster.OnRowExit := OnRowExitFListeMaster;
    FListeSlave := THGrid(GetControl('FLISTESLAVE',true));
    If VH^.PaysLocalisation=CodeISOES then
    Begin
       FListeSlave.ColWidths[6]:=-1 ;
       FListeSlave.ColWidths[7]:=-1 ;
       FListeSlave.ColWidths[8]:=-1 ;
       TFVierge(ECRAN).HMTrad.ResizeDBGrid:=TRUE ;
       TFVierge(ECRAN).HMTrad.ResizeGridColumns(FListeSlave) ;
    End ; //XVI 24/02/2005

    // YMO 27/12/2005 FQ 10364 Colonne "CPT TVA R.G" cachée en environnement non BTP
    {$IFNDEF BTP}
    FListeSlave.ColWidths[8]:=-1 ;
    {$ENDIF}
    if not GereFARFAE then
    begin
      FListeSlave.ColWidths[9] := -1 ;
      FListeSlave.ColWidths[10] := -1 ;
    end;

    FListeSlave.OnRowExit := OnRowExitFListeSlave;
    FListeSlave.OnDblClick := OnDblClickFListeSlave;

    // recup du panel
    PInfo := THPanel(GetControl('PINFO',true));

    // recup des 4 btn de nav
    first := TToolbarbutton97(GetControl('FIRST',true));
    prev := TToolbarbutton97(GetControl('PREV',true));
    next := TToolbarbutton97(GetControl('NEXT',true));
    last := TToolbarbutton97(GetControl('LAST',true));
    first.OnClick := OnClickFirst;
    prev.OnClick := OnClickPrev;
    next.OnClick := OnClickNext;
    last.OnClick := OnClickLast;

    // enable des btn de nav ...
    first.Enabled := false;
    prev.Enabled := false;
    next.Enabled := false;
    last.Enabled := false;

    // recup d'un btn
    BTVAEtab := TToolBarButton97(GetControl('BTVAETAB',true));
    BTVAEtab.OnClick := OnClickBTVAEtab;

    if (TVA) then
    begin // TVA
        FCodeTva := VH^.DefCatTVA;
        Ecran.HelpContext := 1170000;

        SetControlVisible('CC_LIBRE',true);
        SetControlVisible('TCC_LIBRE',true);

        if (ctxMode in V_PGI.PGIContexte) then BTVAEtab.Visible := true;
    end
    else
    begin // TPF
        FCodeTva := VH^.DefCatTPF;
        Ecran.HelpContext := 1175000;

        SetControlVisible('CC_LIBRE',false);
        SetControlVisible('TCC_LIBRE',false);

        BTVAEtab.Visible := false;
    end;
    If VH^.PaysLocalisation=CodeISOES then Begin
       SetControlVisible('CC_LIBRE',FALSE) ;
       SetControlVisible('TCC_LIBRE',FALSE) ;
    End ; //XVI 24/02/2005
    // 1 : Régime
    FlisteSlave.ColEditables[1] := false;
    // 2 : Taux Achat
    FListeSlave.ColTypes[2] := 'R';
    FlisteSlave.ColFormats[2] := '0.000';
    // 3 : Taux Vente
    FListeSlave.ColTypes[3] := 'R';
    FlisteSlave.ColFormats[3] := '0.000';
    // 4 : Compte acaht
    FlisteSlave.ColLengths[4] := GetParamSocSecur('SO_LGCPTEGEN',0);
    // 5 : Compte vente
    FlisteSlave.ColLengths[5] := GetParamSocSecur('SO_LGCPTEGEN',0);
    // 6 : Compte acaht encais
    FlisteSlave.ColLengths[6] := GetParamSocSecur('SO_LGCPTEGEN',0);
    // 7 : compte vente encais
    FlisteSlave.ColLengths[7] := GetParamSocSecur('SO_LGCPTEGEN',0);
    // 8 : compte R.G.
    FlisteSlave.ColLengths[8] := GetParamSocSecur('SO_LGCPTEGEN',0);
    if GereFARFAE then
    begin
      FlisteSlave.ColLengths[9] :=  GetParamSocSecur('SO_LGCPTEGEN',0);
      FlisteSlave.ColLengths[10] :=  GetParamSocSecur('SO_LGCPTEGEN',0);
      LesColsFListeSave := 'TV_REGIME;TV_TAUXACH;TV_TAUXVTE;TV_CPTEACH;TV_CPTEVTE;TV_ENCAISACH;TV_ENCAISVTE;TV_CPTVTERG;TV_CPTACHFARFAE;TV_CPTVTEFARFAE';
    end else
    begin
      LesColsFListeSave := 'TV_REGIME;TV_TAUXACH;TV_TAUXVTE;TV_CPTEACH;TV_CPTEVTE;TV_ENCAISACH;TV_ENCAISVTE;TV_CPTVTERG';
    end;

    // liste des valeur a edité ...
    // recup des regime existant ....
    QLoc := OpenSelect('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="RTV" ORDER BY CC_CODE', FDossier );

    LiRegime := HTStringList.Create;
    LiRegime.Clear;

    while (not QLoc.Eof) do
    begin
        LiRegime.Add(QLoc.Fields[0].AsString);
        QLoc.Next;
    end;

    Ferme(QLoc);
    { FQ 16909 BVE 05.11.07 }
    FListeSlave.Options := FListeSlave.Options + [goTabs];
    { END FQ 16909 }
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTVATPF.OnClose;
begin
    Inherited;

    FreeAndNil(HeaderTOB);
    FreeAndNil(DetailTOB);

    FreeAndNil(TOBGene);

    FreeAndNil(LiRegime);

    AvertirTable('TTREGIMETVA');
    AvertirTable('TTTVAENCAISSEMENT');
    AvertirTable('TTTVA');
    AvertirTable('TTTPF');
    AvertirTable('TTCONTROLETVA');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTVATPF.OnDisplay;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTVATPF.OnCancel;
var
    cancel : boolean;
begin
    Inherited;

    OnUndo := true;

    // reaffichage
    HeaderTOB.Detail[FListeMaster.row-1].PutEcran(ecran,PInfo);
    HeaderTOB.Detail[FListeMaster.row-1].PutLigneGrid(FListeMaster,FListeMaster.row,false,false,'CC_CODE;CC_LIBELLE');

    // if oncreat !!!
    if (OnCreat) then
    begin
        // si un seul enreg
        if (HeaderTOB.Detail.Count = 1) then exit // normalement ca arrive jamais ... le btn est disable
        else
        begin
            FListeMaster.DeleteRow(FListeMaster.row);
            OnCreat := false;

            // affichage du nvl enreg
            HeaderTOB.Detail[FListeMaster.row-1].PutEcran(ecran,PInfo);
            HeaderTOB.Detail[FListeMaster.row-1].PutLigneGrid(FListeMaster,FListeMaster.row,false,false,'CC_CODE;CC_LIBELLE');
            OnRowExitFListeMaster(nil,FListeMaster.row,cancel,true);
        end;
    end;

    // disable des ctrl
    SetControlEnabled('CC_CODE',False);
    SetControlEnabled('bInsert',true);
    SetControlEnabled('bDelete',true);

    OnUndo := false;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPTVATPF.OnRowExitFListeMaster(Sender : TObject ; ou : Longint ; var Cancel : Boolean ; Chg : Boolean);
var
    SQLSlave : string;
    ToUpdate : boolean;
    i : integer;
    tmp : TOB;
begin
    ToUpdate := false;

    // est modifié ??
    if (not (OnRead)) then ToUpdate := (IsModifiedMaster(ou-1) or OnCreat);
    // maj si necessaire
    if (ToUpdate) then
    begin
        if (PGIAsk(TraduireMemoire(Messages[3]),ecran.caption) = mrYes) then cancel := not SaveDBMaster(ou-1)
        else OnCancel;
    end;
    // svg raté ??
    if (Cancel) then exit;

    // disable des ctrl
    first.Enabled := (not (FListeMaster.row = 1));
    prev.Enabled := (not (FListeMaster.Row = 1));
    next.Enabled := (not (FListeMaster.row = (FListeMaster.RowCount - 1)));
    last.Enabled := (not (FListeMaster.row = (FListeMaster.RowCount - 1)));

    // put des valeur sur l'ecran
    if (not (FListeMaster.Row = 0)) then
    begin
        HeaderTOB.Detail[FListeMaster.Row-1].PutEcran(ecran,PInfo);
        FreeAndNil(DetailTob);

        if (OnWillCreat) then
        begin
            // creation des enreg
            DetailTOB := TOB.Create('TXCPTTVA',nil,-1);
            for i:=0 to LiRegime.Count-1 do
            begin
                // creation des ligne de regime
                tmp := TOB.Create('TXCPTTVA',DetailTOB,-1);
                // remplissage
                if (TVA) then tmp.PutValue('TV_TVAOUTPF','TX1')
                else tmp.PutValue('TV_TVAOUTPF','TX2');
                tmp.PutValue('TV_CODETAUX',FListeMaster.Cells[1,FListeMaster.Row]);
                tmp.PutValue('TV_REGIME',LiRegime[i]);
            end;
            OnWillCreat := false;
        end
        else
        begin
            // lecteure des enreg existant ....
            if (TVA) then SQLSlave := 'SELECT * FROM TXCPTTVA WHERE TV_TVAOUTPF="TX1" AND TV_CODETAUX="' + FListeMaster.Cells[1,FListeMaster.Row] + '" ORDER BY TV_CODETAUX,TV_REGIME'
            else SQLSlave := 'SELECT * FROM TXCPTTVA WHERE TV_TVAOUTPF="TX2" AND TV_CODETAUX="' + FListeMaster.Cells[1,FListeMaster.Row] + '" ORDER BY TV_CODETAUX,TV_REGIME';

            DetailTOB := TOB.Create('TXCPTTVA',nil,-1);
            DetailTOB.LoadDetailDBFromSQL('TXCPTTVA',SQLSlave,false,false);

            // ajout des detail ki existe po
            for i:=0 to LiRegime.Count-1 do
            begin
                if (DetailTOB.FindFirst(['TV_REGIME'],[LiRegime[i]],true) = nil ) then
                begin
                    // creation des ligne de regime
                    tmp := TOB.Create('TXCPTTVA',DetailTOB,i);
                    // remplissage
                    if (TVA) then tmp.PutValue('TV_TVAOUTPF','TX1')
                    else tmp.PutValue('TV_TVAOUTPF','TX2');
                    tmp.PutValue('TV_CODETAUX',FListeMaster.Cells[1,FListeMaster.Row]);
                    tmp.PutValue('TV_REGIME',LiRegime[i]);
                    // creation phisique
                    tmp.InsertDB(nil);
                end;
            end;
        end;
        // affichage du detail
        DetailTOB.PutGridDetail(FListeSlave,false,false,LesColsFListeSave,true);
    end;

    if (TVA) then ecran.caption := TraduireMemoire(Messages[0]) + ' ' + FListeMaster.Cells[1,FListeMaster.Row] + ' ' + FListeMaster.Cells[2,FListeMaster.Row]
    else ecran.caption := TraduireMemoire(Messages[1]) + ' ' + FListeMaster.Cells[1,FListeMaster.Row] + ' ' + FListeMaster.Cells[2,FListeMaster.Row];
    UpdateCaption(ecran);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPTVATPF.OnRowExitFListeSlave(Sender : TObject ; ou : Longint ; var Cancel : Boolean ; Chg : Boolean);
var
    ToUpdate : boolean;
begin
    ToUpdate := false;

    // est modifié ??
    if (not (OnRead)) then ToUpdate := IsModifiedSlave(ou-1);
    // maj si necessaire
    if (ToUpdate) then {if (PGIAsk(TraduireMemoire(Messages[3]),ecran.caption) = mrYes) then} cancel := not SaveDBSlave(ou-1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
Procedure TOF_CPTVATPF.OnDblClickFListeSlave(Sender : TObject);
begin
    if (FListeSlave.Col > 3) then
    begin
        LookupList(FListeSlave,Ecran.Caption,'GENERAUX','G_GENERAL','G_LIBELLE','','G_GENERAL',true,1);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPTVATPF.OnClickFirst(Sender : TObject);
begin
    // deplacement dans la grille
    FListeMaster.gotoRow(1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPTVATPF.OnClickPrev(Sender : TObject);
begin
    // deplacement dans la grille
    if (FListeMaster.row - 1 > 0) then FListeMaster.gotoRow(FListeMaster.row - 1)
    else
    begin
        Prev.enabled := false;
        last.enabled := false;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPTVATPF.OnClickNext(Sender : TObject);
begin
    // deplacement dans la grille
    if (FListeMaster.Row + 1 < FListeMaster.RowCount) then FListeMaster.gotoRow(FListeMaster.Row + 1)
    else
    begin
        first.enabled := false;
        next.enabled := false;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPTVATPF.OnClickLast(Sender : TObject);
begin
    // deplacement dans la grille
    FListeMaster.gotoRow(FListeMaster.RowCount - 1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 19/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPTVATPF.OnClickBTVAEtab(Sender : TOBject);
begin
{$IFDEF EAGLCLIENT}
    // TODO
{$ELSE}
    if (not (GetControlText('CC_CODE') = '')) then ParamTVAEtab( GetControlText('CC_CODE'),FCodeTva);
{$ENDIF}
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPTVATPF.IsModifiedMaster(row : integer) : boolean;
begin
    result := true;

    if (not (GetControlText('CC_CODE') = HeaderTOB.Detail[row].GetValue('CC_CODE'))) then exit;
    if (not (GetControlText('CC_LIBELLE') = HeaderTOB.Detail[row].GetValue('CC_LIBELLE'))) then exit;
    if (not (GetControlText('CC_LIBRE') = HeaderTOB.Detail[row].GetValue('CC_LIBRE'))) then exit;

    result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPTVATPF.SaveDBMaster(row : integer) : boolean;
var
    i : integer;
begin
    result := false;

    if (not CanSaveMaster(row)) then exit;

    // recup des valeur dans la TOB
    HeaderTOB.Detail[row].GetEcran(ecran,PInfo);
    HeaderTOB.Detail[row].PutValue('CC_ABREGE',copy(GetControlText('CC_LIBELLE'),1,17));

    // update or insert ....
    if (OnCreat) then
    begin
        result := InsertTobMS( HeaderTOB.Detail[row], FDossier );
        for i:=0 to DetailTOB.Detail.Count-1 do SaveDBSlave(i);
    end
    else
    begin
        result := UpdateTobMS( HeaderTOB.Detail[row], FDossier );
        if (IsModifiedSlave(FListeSlave.Row-1)) then SaveDBSlave(FListeSlave.Row-1);
    end;

    OnCreat := false;

    // recup des valeur sur l'ecran
    HeaderTOB.Detail[row].PutEcran(ecran,PInfo);
    HeaderTOB.Detail[row].PutLigneGrid(FListeMaster,row+1,false,false,'CC_CODE;CC_LIBELLE');

    // disable des ctrl
    SetControlEnabled('CC_CODE',False);
    SetControlEnabled('bInsert',true);
    SetControlEnabled('bDelete',true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 18/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function TOF_CPTVATPF.CanSaveMaster(row : integer) : boolean;
begin
    result := false;

    if (OnCreat) then
    begin
        if (GetControlText('CC_CODE') = '') then
        begin
            PGIBox(Messages[5],Ecran.caption);
            SetFocusControl('CC_CODE');
            exit;
        end;
        if (SiExisteCode) then
        begin
            PGIBox(Messages[7],Ecran.caption);
            SetFocusControl('CC_CODE');
            exit;
        end;
    end;

    if (GetControlText('CC_LIBELLE') = '') then
    begin
        PGIBox(Messages[6],Ecran.caption);
        SetFocusControl('CC_LIBELLE');
        exit;
    end;
    if ((TVA) and (SiExisteIndice)) then
    begin
        PGIBox(Messages[8],Ecran.caption);
        SetFocusControl('CC_LIBRE');
        exit;
    end;
    if (VH^.PaysLocalisation=CodeISOES) and (not ExisteSQL('select 1 from TXCPTTVA where TV_CODETAUX="'+GetControltext('CC_CODE')+'" group BY TV_CODETAUX having Count(TV_CODETAUX)>='+inttostr((FListeSlave.RowCount-FListeSlave.FixedRows)*(1+ord(not oncreat))))) then
       PGIBox(Messages[20],Ecran.caption); //XVI 24/02/2005
    result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPTVATPF.IsModifiedSlave(row : integer) : boolean;
begin
    result := true;

    if (not (FListeSlave.Cells[1,row+1] = DetailTOB.Detail[row].GetValue('TV_REGIME'))) then exit;
    if (not (FListeSlave.Cells[2,row+1] = DetailTOB.Detail[row].GetValue('TV_TAUXACH'))) then exit;
    if (not (FListeSlave.Cells[3,row+1] = DetailTOB.Detail[row].GetValue('TV_TAUXVTE'))) then exit;
    if (not (FListeSlave.Cells[4,row+1] = DetailTOB.Detail[row].GetValue('TV_CPTEACH'))) then exit;
    if (not (FListeSlave.Cells[5,row+1] = DetailTOB.Detail[row].GetValue('TV_CPTEVTE'))) then exit;
    if (not (FListeSlave.Cells[6,row+1] = DetailTOB.Detail[row].GetValue('TV_ENCAISACH'))) then exit;
    if (not (FListeSlave.Cells[7,row+1] = DetailTOB.Detail[row].GetValue('TV_ENCAISVTE'))) then exit;
    if (not (FListeSlave.Cells[8,row+1] = DetailTOB.Detail[row].GetValue('TV_CPTVTERG'))) then exit;
    if GereFARFAE then
    begin
      if (not (FListeSlave.Cells[9,row+1] = DetailTOB.Detail[row].GetValue('TV_CPTACHFARFAE'))) then exit;
      if (not (FListeSlave.Cells[10,row+1] = DetailTOB.Detail[row].GetValue('TV_CPTVTEFARFAE'))) then exit;
    end;
    result := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPTVATPF.SaveDBSlave(row : integer) : boolean;
begin
    result := false;

    if (not CanSaveSlave(row)) then exit;

    // recup des valeur dans la TOB
    DetailTOB.Detail[row].GetLigneGrid(FListeSlave,Row+1,';'+LesColsFListeSave);
    // update or insert ....
    if (OnCreat) then
    begin
        DetailTOB.Detail[row].PutValue('TV_CODETAUX',GetControlTExt('CC_CODE'));
        result := DetailTOB.Detail[row].InsertDB(nil,false)
    end
    else result := DetailTOB.Detail[row].UpdateDB(false,false);

    // recup des valeur sur l'ecran
    DetailTOB.Detail[row].PutLigneGrid(FListeSlave,row+1,false,false,LesColsFListeSave);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 18/06/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CPTVATPF.CanSaveSlave(row : integer) : boolean;
var LeWhere : string;
begin
    result := false;

    if (TOBGene = nil) then
    begin
        TOBGene := TOB.Create('GENERAUX',nil,-1);
{$IFDEF GCGC}
        { En GC, on autorise les comptes de charges et produits pour la TPF }
        if not TVA then
          LeWhere := 'SELECT G_GENERAL FROM GENERAUX WHERE G_COLLECTIF="-" AND G_NATUREGENE NOT IN ("COC","COD","COF","COS","IMO")'
          else
          LeWhere := 'SELECT G_GENERAL FROM GENERAUX WHERE G_COLLECTIF="-" AND G_NATUREGENE NOT IN ("COC","COD","COF","COS","IMO","CHA","PRO")';
{$ELSE GCGC}
        LeWhere := 'SELECT G_GENERAL FROM GENERAUX WHERE G_COLLECTIF="-" AND G_NATUREGENE NOT IN ("COC","COD","COF","COS","IMO","CHA","PRO")';
{$ENDIF GCGC}
        TOBGene.LoadDetailFromSQL(LeWhere, false, false);
    end;

    if ((StrToFloat(FListeSlave.Cells[2,row+1]) <> 0.00) and (FListeSlave.Cells[4,row+1] = '')) then
    begin
        PGIBox(Messages[11],Ecran.caption);
        FListeSlave.Col := 4;
        FListeSlave.SetFocus;
        exit;
    end;

    if ((StrToFloat(FListeSlave.Cells[3,row+1]) <> 0.00) and (FListeSlave.Cells[5,row+1] = '')) then
    begin
        PGIBox(Messages[12],Ecran.caption);
        FListeSlave.Col := 5;
        FListeSlave.SetFocus;
        exit;
    end;

    // autre ...
    if (FListeSlave.Cells[4,row+1] <> '') then
    begin
        if (TOBGene.FindFirst(['G_GENERAL'],[FListeSlave.Cells[4,row+1]],true) = nil) then
        begin
            PGIBox(Messages[15],Ecran.caption);
            FListeSlave.Col := 4;
            FListeSlave.SetFocus;
            exit;
        end;
    end;

    if (FListeSlave.Cells[5,row+1] <> '') then
    begin
        if (TOBGene.FindFirst(['G_GENERAL'],[FListeSlave.Cells[5,row+1]],true) = nil) then
        begin
            PGIBox(Messages[16],Ecran.caption);
            FListeSlave.Col := 5;
            FListeSlave.SetFocus;
            exit;
        end;
    end;

    if (FListeSlave.Cells[6,row+1] <> '') then
    begin
        if (TOBGene.FindFirst(['G_GENERAL'],[FListeSlave.Cells[6,row+1]],true) = nil) then
        begin
            PGIBox(Messages[17],Ecran.caption);
            FListeSlave.Col := 6;
            FListeSlave.SetFocus;
            exit;
        end;
    end;

    if (FListeSlave.Cells[7,row+1] <> '') then
    begin
        if (TOBGene.FindFirst(['G_GENERAL'],[FListeSlave.Cells[7,row+1]],true) = nil) then
        begin
            PGIBox(Messages[18],Ecran.caption);
            FListeSlave.Col := 7;
            FListeSlave.SetFocus;
            exit;
        end;
    end;

    result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 18/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
function TOF_CPTVATPF.SiExisteCode : Boolean;
var
    Q : TQuery ;
begin
    if (TVA) then Q := OpenSelect('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="TX1" AND CC_CODE="' + GetControlText('CC_CODE') + '"',FDossier)
    else Q := OpenSelect('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="TX2" AND CC_CODE="' + GetControlText('CC_CODE') + '"',FDossier);
    result := (not Q.Eof);
    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 18/06/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
function TOF_CPTVATPF.SiExisteIndice : boolean;
var
    Q : TQuery;
begin
    if (Valeuri(GetControlText('CC_LIBRE')) > 0) then
    begin
        Q := OpenSelect('SELECT COUNT(*) FROM CHOIXCOD WHERE CC_TYPE="TX1" AND CC_CODE<>"' + GetControlText('CC_CODE') + '" AND CC_LIBRE="' + GetControlText('CC_LIBRE') + '"',FDossier);
        if ((not Q.EOF) and (Q.Fields[0].AsInteger > 0)) then result := true
        else result := false;
        Ferme(Q);
    end
    else Result := false;
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPTVATPF]);
end.
