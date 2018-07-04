{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/01/2018
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISTRFPOC ()
Mots clefs ... : TOF;BTSAISTRFPOC
*****************************************************************}
Unit BTSAISTRFPOC_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     uDbxDataSet,
     mul,
     Fe_Main,

{$else}
     eMul,
{$ENDIF}
     Menus,
     uTob,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     Vierge,
     HSysMenu,
     UTOF,
     UTransferts,
     AglInit ;

const
  TheChampsFrom = 'ECCC;ARTSAISIE;SOUSTRAITANT;CODEMARCHE;LIBELLE;QTE;UNITE;PA;QTEINIT;MONTANTINIT;MONTANT';
  TheChampsDest = 'ECCC;ARTSAISIE;SOUSTRAITANT;CODEMARCHE;LIBELLE;QTE;UNITE;PA;MONTANT';

Type
  TmodeFoncTRF = (TmfCreation,TmfModif);
  
  TTypeGridSaisie = (TTgsFrom,TTgsDest);

  TOF_BTSAISTRFPOC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    XX : TForm;
    ModeFonc : TmodeFoncTRF;
    ThePop : TpopupMenu;
    MtInit,MtTransf,MtEcart : double;
    SuiteWhere : string;
    TOBNewLigne : TOB;
    TOBOUvrages : TOB;
    TOBTRFPOC,TOBParag : TOB;
    TOBFROM : TOB;
    TOBDEST : TOB;
    TOBPhases : TOB;
    TOBPiece : TOB;
    GSFROM,GSDEST : THGrid;
    BSELDEST : TToolbarButton97;
    StCellFrom,StCellDest : string;
    GSART,GSLIB,GSQTE,GSPU,GSMARCHE,GSMONTANTD,GSMONTANTF : Integer;
    procedure CreateTOB;
    procedure GetControls;
    procedure DefiniLibelleDestination;
    procedure DefiniTOBSaisi;
    procedure AddChampsSupSaisie(TT : TOB);
    procedure RenseigneEcran;
    procedure ConstitueGrillesSaisie;
    procedure DefinieGrilleSaisie(GS: THgrid; TypeGrille : TTypeGridSaisie);
    procedure LibereTobs;
    procedure ChargeLesParags;
    procedure AfficheGrille(TOBT: TOB; GS: THGrid; TypeGrid : TTypeGridSaisie);
    procedure DefiniPied;
    procedure SetEvents (Status : boolean);
    procedure SetEventForm(Status : boolean);
    procedure SetEventGrid(GS: THGrid; TypeGrid : TTypeGridSaisie;Status :boolean);
    procedure BSELDestClick (Sender : TObject);
    //
    procedure GSCellEnterFrom(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExitFrom(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellEnterDest(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExitDest(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSRowEnterDest(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExitDest(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    function IsLigneComplete(Arow: Integer): Boolean;
    function AjouteDetail: TOB;
    function IsExitInterdit(Acol, Arow: integer): boolean;
    //
    procedure TraiteCodeArt (ACol,Arow : Integer; var Cancel : boolean);
    procedure TraiteLibelle (ACol,Arow : Integer; var Cancel : boolean);
    procedure TraiteQteFrom (ACol,Arow : Integer; var Cancel : boolean);
    procedure TraiteQte (ACol,Arow : Integer; var Cancel : boolean);
    procedure TraitePu (ACol,Arow : Integer; var Cancel : boolean);
    procedure TraiteMontantDest (ACol,Arow : Integer; var Cancel : boolean);
    procedure TraiteMontantFrom (ACol,Arow : Integer; var Cancel : boolean);
    procedure TraiteMarche (ACol,Arow : Integer; var Cancel : boolean);
    //
    procedure CalculeLaLigne(TOBL : TOB; TypeGrid: TTypeGridSaisie);
    procedure AjouteLaLigne (TOBL : TOB; TypeGrid : TTypeGridSaisie; Montant : double; Sens : string ='+');
    procedure AffichelaLigne(TOBL: TOB; TypeGrid: TTypeGridSaisie;Ligne: integer);
    procedure GSElipsisClick(Sender: TObject);
    procedure SetArticleToLigne(TOBX,TOBA : TOB);
    function FindArticle(CodeArticle: string;var TOBA: TOB): Boolean;
    function RechercheArt(Arow: integer): boolean;
    procedure ZoneSuivanteOuOk(GS: THgrid; var ACol, ARow: Integer;var Cancel: boolean);
    procedure BeforePopup (Sender : Tobject);
    procedure AffectInterneClick (sender : TObject);
    procedure AffecteST (Sender : TObject);
    function FindCodeArticle(CodeArticle: string; var TOBA: TOB): Boolean;
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SupLigneClick;
    procedure BeforeInsertTOB;
    procedure AjouteLesDatas;
    procedure DelChampsSupSaisie(TT: TOB);
    procedure AfterInsertTOB;
    procedure AjouteDetailOuv(TOBL : TOB) ;
    procedure DupliqueDetailOuv(TOBL, TOBNL: TOB);
    procedure BdeleteClick (Sender : TObject);
    procedure ConstituePhaseInterne;
    procedure RecalculeLigneFromOuvrage(TOBL, TOBDEPART: TOB);
    procedure RecalculeLesOuvrages(TOBRECALC: TOB);

end ;

Implementation

uses Grids,SelectPhase,Facture,FactTOB,UtilArticle,FactArticle,FactUtil,FactFormule,
     FactPiece,FactCalc,Windows,Messages,FactComm,FactOuvrage,uEntCommun,
     NomenUtil,FactureBTP,UtilPGI,UCotraitance,USpecifPOC,EntGC;

procedure TOF_BTSAISTRFPOC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISTRFPOC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISTRFPOC.OnUpdate ;
begin
  Inherited ;
  if ARRONDI(MtEcart,2)<> 0 then
  begin
    PGIInfo('IMPOSSIBLE : Le montant reste n''est pas nul');
    TFVierge(ecran).ModalResult := 0;
    Exit;
  end;
  if TOBTRFPOC.GetString ('BT2_PHASETRA') = '' then
  begin
    PGIInfo('IMPOSSIBLE : Vous devez sélectionner une destination');
    TFVierge(ecran).ModalResult := 0;
    Exit;
  end;
  BeforeInsertTOB;
  AjouteLesDatas;
  AfterInsertTOB;
  NumeroteLignesGC(TFFacture(XX).GS,TOBPiece,false,true);
  TOBTRFPOC.SetString('OKOK','X');
end ;

procedure TOF_BTSAISTRFPOC.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISTRFPOC.OnArgument (S : String ) ;
var Acol,Arow : Integer;
    cancel : Boolean;
begin
  Inherited ;
  XX := nil;
  if CurrentTransfert <> nil then
  begin
    XX := CurrentTransfert.CurrentSaisie as TFFacture;
  end;
  TOBTRFPOC := LaTOB;
  if TOBTRFPOC.getString('MODE')= 'CREATION' then ModeFonc := TmfCreation
                                             else ModeFonc := TmfModif;
  TOBPIece := TFFacture (XX).LaPieceCourante;
  TOBOUvrages := TFFacture(XX).TheTOBOuvrage;
  SuiteWhere := FabricWhereNatArt(TOBPiece.getString('GP_NATUREPIECEG'), TOBPiece.getString('GP_DOMAINE'),'');
  //
  CreateTOB;
  ChargeLesParags;
  GetControls;
  DefiniTOBSaisi;
  ConstitueGrillesSaisie;
  //
  RenseigneEcran;
  //
  SetEvents (true);
  GSFROM.col := GSMONTANTF;
  Acol := GSFROM.Col;
  Arow := 1;
  cancel := False;
  GSCellEnterFrom(Self,Acol,Arow,cancel);
  //
  GSRowEnterDest(Self,1,cancel,false);
  Acol := GSART;
  Arow := 1;
  GSCellEnterDest(Self,Acol,Arow,cancel);
end ;

procedure TOF_BTSAISTRFPOC.OnClose ;
begin
  LibereTobs;
  Inherited ;
end ;

procedure TOF_BTSAISTRFPOC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISTRFPOC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISTRFPOC.CreateTOB;
begin
  TOBFROM := TOB.create ('LES PROV',nil,-1);
  TOBDEST := TOB.create ('LES DEST',nil,-1);
  TOBParag := TOB.Create ('LES PARAG',nil,-1);
  TOBNewLigne := TOB.Create ('LES NEW LIG',nil,-1);
  TOBPhases := TOB.Create('LES PHASES',nil,-1);
end;

procedure TOF_BTSAISTRFPOC.GetControls;
begin
  GSFROM := THgrid(GetControl('GSFROM'));
  GSDEST := THgrid(GetControl('GSDEST'));
  BSELDEST := TToolbarButton97(GetControl('BSELDEST'));
  ThePop := TPopupMenu(GetControl('POPSEL'));
end;

procedure TOF_BTSAISTRFPOC.DefiniTOBSaisi;
var II : Integer;
    TF,TD,TL,TOBL : TOB;
begin
  MtInit := 0;
  MtTransf := 0;
  for II := 0 to TOBTRFPOC.detail.count -1 do
  begin
    if TOBTRFPOC.detail[II].GetString('BT3_TYPELIGNETRF')='000' then
    begin
      TF := TOB.Create ('BTRFDETAIL',TOBFROM,-1);
      TF.Dupliquer(TOBTRFPOC.detail[II],False,true);
      TF.Data := TOBTRFPOC.detail[II].Data;
      AddChampsSupSaisie(TF);
      MtInit := MtInit + TF.GetDouble('MONTANT');
    end else
    begin
      if TOBTRFPOC.detail[II].GetBoolean('BT3_CONTREP') then Continue; 
      TD := TOB.Create ('BTRFDETAIL',TOBDEST,-1);
      TD.Dupliquer(TOBTRFPOC.detail[II],False,true);
      TOBL := TOB(TOBTRFPOC.detail[II].Data);
      TL := NewTOBLigne(TOBNewLigne,0);
      TL.Dupliquer(TOBL,false,true);
      TD.Data := TL;
      AddChampsSupSaisie(TD);
      MtTransf := MtTransf + TD.GetDouble('MONTANT');
    end;
  end;
  MtEcart := MtInit - MtTransf;
  if TOBDEST.Detail.count = 0 then
  begin
    TD := TOB.Create ('BTRFDETAIL',TOBDEST,-1);
    AddChampsSupSaisie(TD);
  end;
end;

procedure TOF_BTSAISTRFPOC.AddChampsSupSaisie(TT : TOB);
var TOBL,TOBLOUV,TOBO  : TOB;
    MontantInit,QTeInit,MtDejaTransf: double;
begin
  MontantInit := 0;
  QTeInit := 0;
  TOBL := TOB(TT.Data);
  if TOBL <> nil then
  begin
    if TOBL.GetString('GL_TYPELIGNE')='SD' then
    begin
      MtDejaTransf := 0;
      TOBO := FindLigneOuvFromUnique (TOBL,TOBOUvrages,TOBL.GetInteger('UNIQUEBLO'));
      if TOBO <> nil then
      begin
        TOBLOUV := FindLigneDocFromDetail (TOBPiece,TOBL);
        if TOBLOUV <> nil then
        begin
          MontantInit := arrondi(TOBO.GetDouble('BLO_MONTANTPA') * Abs(TOBLOUV.GetDouble('GL_QTEFACT')),TFFacture(XX).DEV.decimale);
          QteInit := arrondi(MontantInit/TOBL.GetDouble('GL_DPA'),V_PGI.okdecQ);
        end;
      end;
    end else
    begin
      MontantINit := TOBL.GetDouble('GL_MONTANTPA');
      QteInit :=  TOBL.GetDouble('GL_QTEFACT');
    end;
  end;
  if not TT.FieldExists('ARTSAISIE') then TT.AddChampSupValeur('ARTSAISIE','') else TT.SetString('ARTSAISIE','');
  if not TT.FieldExists('ARTICLE') then TT.AddChampSupValeur('ARTICLE','') else TT.SetString('ARTICLE','');
  if not TT.FieldExists('LIBELLE') then TT.AddChampSupValeur('LIBELLE','') else TT.SetString('LIBELLE','');
  if not TT.FieldExists('UNITE') then TT.AddChampSupValeur('UNITE','') else TT.SetString('UNITE','');
  if not TT.FieldExists('QTE') then TT.AddChampSupValeur('QTE',0) else TT.SetDouble('QTE',0);
  if not TT.FieldExists('PA') then TT.AddChampSupValeur('PA',0) else TT.SetDouble('PA',0);
  if not TT.FieldExists('QTEINIT') then TT.AddChampSupValeur('QTEINIT ',0) else TT.SetDouble('QTEINIT',0);
  if not TT.FieldExists('MONTANTINIT') then TT.AddChampSupValeur('MONTANTINIT',0) else TT.SetDouble('MONTANTINIT',0);
  if not TT.FieldExists('MONTANT') then TT.AddChampSupValeur('MONTANT',0) else TT.SetDouble('MONTANT',0);
  if not TT.FieldExists('SOUSTRAITANT') then TT.AddChampSupValeur('SOUSTRAITANT','') else TT.SetString('SOUSTRAITANT','');
  if not TT.FieldExists('CODEMARCHE') then TT.AddChampSupValeur('CODEMARCHE','') else TT.SetString('CODEMARCHE','');
  if TOBL <> nil then
  begin
    TT.SetString('ARTSAISIE',TOBL.GetString('GL_CODEARTICLE'));
    TT.SetString('ARTICLE',TOBL.GetString('GL_ARTICLE'));
    TT.SetString('LIBELLE',TOBL.GetString('GL_LIBELLE'));
    TT.SetString('UNITE',TOBL.GetString('GL_QUALIFQTEVTE'));
    TT.SetDouble('PA',TOBL.GetDouble('GL_DPA'));
    if TOBL.GetDouble('MTTRANSFERT') <> 0 then
    begin
      if ModeFonc = TmfCreation then
      begin
        TT.SetDouble('MONTANTINIT',MontantInit - TOBL.GetDouble('MTTRANSFERT'));
        TT.SetDouble('QTEINIT',ARRONDI(TT.GetDouble('MONTANTINIT')/TT.GetDouble('PA'),V_PGI.OkDecQ));
        TT.SetDouble('MONTANT',TT.GetDouble('MONTANTINIT'));
        TT.SetDouble('QTE',ARRONDI(TT.GetDouble('MONTANT')/TT.GetDouble('PA'),V_PGI.OkDecQ));
      end else
      begin
        TT.SetDouble('MONTANTINIT',MontantInit - TOBL.GetDouble('MTTRANSFERT') +TT.GetDouble('BT3_MTTRANSFERT'));
        TT.SetDouble('QTEINIT',ARRONDI(TT.GetDouble('MONTANTINIT')/TT.GetDouble('PA'),V_PGI.OkDecQ));
        TT.SetDouble('MONTANT',TT.GetDouble('BT3_MTTRANSFERT'));
        TT.SetDouble('QTE',ARRONDI(TT.GetDouble('MONTANT')/TT.GetDouble('PA'),V_PGI.OkDecQ));
      end;
    end else
    begin
      TT.SetDouble('MONTANTINIT',MontantInit);
      TT.SetDouble('MONTANT',MontantINit);
      TT.SetDouble('QTE',QteInit);
      TT.SetDouble('QTEINIT',QteINit);
    end;
    TT.SetString('SOUSTRAITANT',TOBL.GetString('GL_FOURNISSEUR'));
    TT.SetString('CODEMARCHE',TOBL.GetString('GL_CODEMARCHE'));
  end;
end;

procedure TOF_BTSAISTRFPOC.DelChampsSupSaisie(TT : TOB);
begin
  if not TT.FieldExists('ARTSAISIE') then TT.DelChampSup('ARTSAISIE',true);
  if not TT.FieldExists('ARTICLE') then TT.DelChampSup('ARTICLE',true);
  if not TT.FieldExists('LIBELLE') then TT.DelChampSup('LIBELLE',true);
  if not TT.FieldExists('UNITE') then TT.DelChampSup('UNITE',true);
  if not TT.FieldExists('QTE') then TT.DelChampSup('QTE',true);
  if not TT.FieldExists('QTEINIT') then TT.DelChampSup('QTEINIT',true);
  if not TT.FieldExists('PA') then TT.DelChampSup('PA',true);
  if not TT.FieldExists('MONTANTINIT') then TT.DelChampSup('MONTANTINIT',true);
  if not TT.FieldExists('MONTANT') then TT.DelChampSup('MONTANT',true);
  if not TT.FieldExists('SOUSTRAITANT') then TT.DelChampSup('SOUSTRAITANT',true);
  if not TT.FieldExists('CODEMARCHE') then TT.DelChampSup('CODEMARCHE',true);
end;

procedure TOF_BTSAISTRFPOC.RenseigneEcran;
begin
  DefiniLibelleDestination;
  AfficheGrille(TOBFROM ,GSFROM, TTgsFrom );
  AfficheGrille(TOBDEST ,GSDEST, TTgsDest );
  THSystemMenu(GetCOntrol('HMTrad')).ResizeGridColumns(GSFROM);
  THSystemMenu(GetCOntrol('HMTrad')).ResizeGridColumns(GSDEST);
  DefiniPied;
end;

procedure TOF_BTSAISTRFPOC.DefinieGrilleSaisie (GS : THgrid; TypeGrille : TTypeGridSaisie);

  function WitdhSize (NbCar : integer) : integer;
  begin
    result := GS.Canvas.TextWidth('W') * NbCar;
  end;
  
var FFM,FFQ,FFP : string;
    II : Integer;
begin
  FFM := '#';
  if V_PGI.OkDecV > 0 then
  begin
    FFM := '0.';
    for II := 1 to V_PGI.OkDecV - 1 do
    begin
      FFM := FFM + '#';
    end;
    FFM := FFM + '0';
  end;
  FFP := '#';
  if V_PGI.OkDecP > 0 then
  begin
    FFP := '0.';
    for II := 1 to V_PGI.OkDecP - 1 do
    begin
      FFP := FFP + '#';
    end;
    FFP := FFP + '0';
  end;
  FFQ := '#';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQ := '0.';
    for II := 1 to V_PGI.OkDecQ - 1 do
    begin
      FFQ := FFQ + '#';
    end;
    FFQ := FFQ + '0';
  end;

  GS.VidePile(false);
  //
  if TypeGrille = TTgsDest then GS.ColCount := 9 else GS.colCount := 11;
  //
  GS.Cells[0,0] := '';
  GS.ColWidths[0] := 15;
  GS.ColLengths[0] := 3;
  GS.ColWidths[0] := 10;
  GS.ColAligns[0] := taRightJustify;
  GS.ColEditables[0] := false;
  //
  GS.Cells[1,0] := 'Code';
  GS.ColWidths[1] := WitdhSize(18);
  GS.ColLengths[1] := 18;
  GS.ColAligns[1] := taLeftJustify;
  GS.ColEditables[1] := (TypeGrille = TTgsDest) ;
  GSART := 1;
  //
  GS.Cells[2,0] := 'Sous Trait';
  GS.ColWidths[2] := WitdhSize(12);
  GS.ColLengths[2] := 12;
  GS.ColAligns[2] := taLeftJustify;
  GS.ColEditables[2] := false ;
  //
  GS.Cells[3,0] := 'Code Marché';
  GS.ColWidths[3] := WitdhSize(15);
  GS.ColLengths[3] := 15;
  GS.ColAligns[3] := taLeftJustify;
  GS.ColEditables[3] := false ;
  GSMARCHE := 3;
  //
  GS.Cells[4,0] := 'Libellé';
  GS.ColWidths[4] := WitdhSize(50);
  GS.ColLengths[4] := 70;
  GS.ColAligns[4] := taLeftJustify;
  GS.ColEditables[4] := false;
  GS.ColEditables[4] := (TypeGrille = TTgsDest) ;
  GSLIB := 4;
  //
  GS.Cells[5,0] := 'Qté';
  GS.ColWidths[5] := WitdhSize(12);
  GS.ColLengths[5] := 14;
  GS.ColAligns[5] := taRightJustify;
  GS.ColTypes [5] := 'R';
  GS.ColFormats[5] := FFQ+';'+FFQ+'; ;';
  GS.ColEditables[5] := true ;
  GSQTE := 5;
  //
  GS.Cells[6,0] := 'UN';
  GS.ColWidths[6] := WitdhSize(3);
  GS.ColLengths[6] := 3;
  GS.ColAligns[6] := taCenter;
  GS.ColEditables[6] := false;
  //
  GS.Cells[7,0] := 'Prix Unitaire';
  GS.ColWidths[7] := WitdhSize(10);
  GS.ColLengths[7] := 10;
  GS.ColTypes [7] := 'R';
  GS.ColAligns[7] := taRightJustify;
  GS.ColFormats[7] := FFP+';'+FFP+'; ;';
  GS.ColEditables[7] := (TypeGrille = TTgsDest) ;
  GSPU := 7;
  //
  if TypeGrille = TTgsFrom then
  begin
    GS.Cells[8,0] := 'Qte Initial';
    GS.ColWidths[8] := WitdhSize(12);
    GS.ColLengths[8] := 12;
    GS.ColTypes [8] := 'R';
    GS.ColAligns[8] := taRightJustify;
    GS.ColFormats[8] := FFM+';'+FFM+'; ;';
    GS.ColEditables[8] := false ;
    //
    GS.Cells[9,0] := 'MT Initial';
    GS.ColWidths[9] := WitdhSize(12);
    GS.ColLengths[9] := 12;
    GS.ColTypes [9] := 'R';
    GS.ColAligns[9] := taRightJustify;
    GS.ColFormats[9] := FFM+';'+FFM+'; ;';
    GS.ColEditables[9] := false ;
    //
    GS.Cells[10,0] := 'MT Transfert';
    GS.ColWidths[10] := WitdhSize(12);
    GS.ColLengths[10] := 12;
    GS.ColTypes [10] := 'R';
    GS.ColAligns[10] := taRightJustify;
    GS.ColFormats[10] := FFM+';'+FFM+'; ;';
    GS.ColEditables[10] := true;
    GSMONTANTF := 10;
  end else
  begin
    GS.Cells[8,0] := 'MT Transféré';
    GS.ColWidths[8] := WitdhSize(12);
    GS.ColLengths[8] := 12;
    GS.ColTypes [8] := 'R';
    GS.ColAligns[8] := taRightJustify;
    GS.ColFormats[8] := FFM+';'+FFM+'; ;';
    GS.ColEditables[8] := true;
    GSMONTANTD := 8;
  end;
end;

procedure TOF_BTSAISTRFPOC.ConstitueGrillesSaisie;
begin
  DefinieGrilleSaisie(GSFROM,TTgsFrom);
  DefinieGrilleSaisie(GSDEST,TTgsDest);
end;

procedure TOF_BTSAISTRFPOC.DefiniLibelleDestination;

  function MonteUnCran (Code : string) : string;
  var II,posX : Integer;
  begin
    Result := '';
    PosX := -1;
    for II := Length(Code) downto 0 do
    begin
      if Copy(Code,II,1)='.' then
      begin
        PosX := II;
        break;
      end;
    end;
    if posX > 0 then
    begin
      Result := Copy(Code,1,PosX-1);
    end;
  end;

  function GetLibPhase (Code : string) : string;
  var II : Integer;
  begin
    For II := 0 to TOBParag.detail.count -1 do
    begin
      if TOBParag.detail[II].GetString('BPC_PHASETRA')=Code then
      begin
        Result := TOBParag.detail[II].GetString('BPC_LIBELLE')
      end;
    end;
  end;

var Code,result : string;
begin
  THLabel(GetControl('LNOMDEST')).Caption := 'Non défini';
  if TOBTRFPOC.GetString('BT2_PHASETRA')='' then Exit;
  Code := TOBTRFPOC.GetString('BT2_PHASETRA');
  result := '';
  repeat
    if result <> '' then result := GetLibPhase(Code) + ' - ' +result
                    else result := GetLibPhase(Code);
    Code := MonteUnCran(Code);
  until Code = '';
  if result <> '' then
  begin
    THLabel(GetControl('LNOMDEST')).Caption := result;
  end else
  begin
    THLabel(GetControl('LNOMDEST')).Caption := 'Non défini';
  end;
end;

procedure TOF_BTSAISTRFPOC.LibereTobs;
begin
  TOBFROM.free;
  TOBDEST.free;
  TOBParag.free;
  TOBNewLigne.Free;
  TOBPhases.free;
end;

procedure TOF_BTSAISTRFPOC.ChargeLesParags;
var Affaire : string;
    QQ : TQuery;
begin
  TOBParag.ClearDetail;
  Affaire := TOBPiece.GetString('GP_AFFAIRE');
  if Affaire = '' then Exit;
  QQ := OpenSQL('SELECT * FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+Affaire+'"',True,-1,'',true);
  if Not QQ.eof then
  begin
    TOBParag.LoadDetailDB('PHASESCHANTIER','','',QQ,false);
  end;
  Ferme(QQ);
end;

procedure TOF_BTSAISTRFPOC.AffichelaLigne (TOBL  : TOB; TypeGrid : TTypeGridSaisie; Ligne : integer );
begin
  if TypeGrid = TTgsFrom then TOBL.PutLigneGrid(GSFROM,Ligne,False,False,TheChampsFrom)
                         else TOBL.PutLigneGrid(GSDEST,Ligne,False,False,TheChampsDest);
end;

procedure TOF_BTSAISTRFPOC.AfficheGrille (TOBT  : TOB; GS : THGrid; TypeGrid : TTypeGridSaisie);
var II : Integer;
begin
  if TypeGrid = TTgsFrom then GS.RowCount := TOBT.detail.count +1
                         else GS.RowCount := TOBT.detail.count +2;
  for II := 0 to TOBT.detail.count -1 do
  begin
    AffichelaLigne(TOBT.detail[II],TypeGrid,II+1);
  end;
end;

procedure TOF_BTSAISTRFPOC.DefiniPied;
begin
  MtEcart := MtInit - MtTransf;
  THNumEdit(GetControl('MTINIT')).Value := MtInit;
  THNumEdit(GetControl('MTRESTE')).Value := MtEcart;
end;

procedure TOF_BTSAISTRFPOC.SetEvents(Status: boolean);
begin
  SetEventForm(Status);
  SetEventGrid(GSFROM,TTgsFrom ,Status);
  SetEventGrid(GSDEST,TTgsDest ,Status);
end;

procedure TOF_BTSAISTRFPOC.SetEventForm(Status: boolean);
begin
  if Status then
  begin
    BSELDEST.OnClick := BSELDestClick;
    ThePop.OnPopup := BeforePopup;
    TMenuItem(GetControl('MnAffectST')).OnClick := AffecteST;
    TToolbarButton97 (GetControl('BdeleteALL')).onclick := BdeleteClick;
  end else
  begin
    BSELDEST.OnClick := nil;
    ThePop.OnPopup := nil;
    TMenuItem(GetControl('MnAffectST')).OnClick := nil;
    TToolbarButton97 (GetControl('BdeleteALL')).onclick := nil;
  end;
end;

procedure TOF_BTSAISTRFPOC.SetEventGrid(GS: THGrid;TypeGrid: TTypeGridSaisie; Status: boolean);
begin
  if Status then
  begin
    if Typegrid = TTgsFrom then
    begin
      GS.OnCellEnter := GSCellEnterFrom;
      GS.OnCellExit := GSCellexitFrom;
    end else
    begin
      GS.OnRowEnter := GSRowEnterDest;
      GS.OnRowExit := GSRowExitDest;
      GS.OnCellEnter := GSCellEnterDest;
      GS.OnCellExit := GSCellExitDest;
      GS.OnElipsisClick := GSElipsisClick;
      GS.OnKeyDown := GSKeyDown;
    end;
  end else
  begin
    GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
    GS.OnCellEnter := nil;
    GS.OnCellExit := Nil;
    GS.OnElipsisClick := nil;
    GS.OnKeyDown := nil;
  end;
end;

function TOF_BTSAISTRFPOC.FindCodeArticle(CodeArticle: string;var TOBA : TOB): Boolean;
var QQ : TQuery;
begin
  Result := false;
  TOBA := TFFacture(XX).TheTOBArticles.FindFirst(['GA_CODEARTICLE'],[CodeArticle],true);
  if TOBA <> nil then   BEGIN Result := True; Exit; end;

  if TOBA = nil then
  begin
    TOBA := CreerTOBArt(TFFacture(XX).TheTOBArticles);
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                   'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                   'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_CODEARTICLE="'+CodeArticle+'"',true,-1, '', True);
    if Not QQ.EOF then
    begin
      ChargerTobArt(TOBA,nil,'VEN','',QQ) ;
      Result := True;
    end;
    Ferme(QQ);
  end;
end;

function TOF_BTSAISTRFPOC.FindArticle(CodeArticle: string;var TOBA : TOB): Boolean;
var QQ : TQuery;
begin
  Result := false;
  TOBA := TFFacture(XX).TheTOBArticles.FindFirst(['GA_ARTICLE'],[CodeArticle],true);
  if TOBA <> nil then   BEGIN Result := True; Exit; end;

  if TOBA = nil then
  begin
    TOBA := CreerTOBArt(TFFacture(XX).TheTOBArticles);
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                   'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                   'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+CodeArticle+'"',true,-1, '', True);
    if Not QQ.EOF then
    begin
      ChargerTobArt(TOBA,nil,'VEN','',QQ) ;
      Result := True;
    end;
    Ferme(QQ);
  end;
end;

function TOF_BTSAISTRFPOC.RechercheArt(Arow : integer) : boolean;
var RefArticle,StChamps : string;
    TOBA,TOBL : TOB;
begin
  Result := false;
  TOBL := TOBDEST.Detail[Arow-1];
  TOBA := nil;

  if GSDEST.Cells[GSART,ARow] <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(GSDEST.Cells[GSART,ARow], 1, 18));

  if SuiteWhere <> '' then
  begin
    if StChamps <> '' then StChamps := StChamps + ';XX_WHERE= AND (' + SuiteWhere+')'
                      else StChamps := 'XX_WHERE= AND (' + SuiteWhere+')';
  end;


  RefArticle := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps+';RECHERCHEARTICLE');
  if RefArticle <> '' then
  begin
    result := FindArticle(RefArticle,TOBA);
    if result then
    begin
      GSDEST.Cells[GSART,Arow] := TOBA.GetString('GA_CODEARTICLE');
      SetArticleToLigne(TOBL,TOBA);
    end;
  end;
end;

procedure TOF_BTSAISTRFPOC.GSElipsisClick(Sender: TObject);
var TOBL : TOB;
begin
  if GSDEST.col = GSART then
  begin
    TOBL := TOBDEST.Detail[GSDEST.row-1];  if TOBL.GetString('ARTSAISIE')<> '' then Exit;
    if RechercheArt(GSDEST.row) then
    begin
      CalculeLaLigne(TOBL,TTgsDest);
      AjouteLaLigne (TOBL,TTgsDest,TOBL.GetDouble('MONTANT'));
      AfficheLaLIgne(TOBL,TTgsDest ,GSDEST.row);
      StCellDest := GSDEST.Cells[GSDEST.col,GSDEST.row];
      DefiniPied;
    end;
  end;
end;

procedure TOF_BTSAISTRFPOC.GSCellEnterDest(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
//  ZoneSuivanteOuOk (GSDEST,Acol,Arow,Cancel);
  if cancel then exit;
  GSDEST.ElipsisButton := (GSDEST.Col = GSART) OR (GSDEST.Col = GSMARCHE);
  StCellDest := GSDEST.Cells[GSDEST.Col,GSDEST.Row];
  GSDEST.RowCount := TOBDEST.Detail.count +2;
end;

procedure TOF_BTSAISTRFPOC.GSCellEnterFrom(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteOuOk (GSFROM,Acol,Arow,Cancel);
  if cancel then exit;
  StCellFrom := GSFROM.Cells[GSFROM.Col,GSFROM.Row];
end;

procedure TOF_BTSAISTRFPOC.GSCellExitDest(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
  function ZoneCalc (Acol : Integer) : Boolean;
  begin
    Result := ((Acol = GSART) or (Acol = GSQTE) or (Acol=GSPU) OR (Acol=GSMONTANTD));
  end;
var OldMt : double;
    TOBL : TOB;
begin
  Cancel := IsExitInterdit(Acol,Arow);
  if cancel then exit;
  if GSDEST.Cells[ACol,Arow] = StCellDest then exit;
  TOBL := TOBDEST.detail[Arow-1];
  OldMt := TOBL.GetDouble('MONTANT');
  if Acol = GSART then // Code Article
  begin
    TraiteCodeArt (ACol,Arow,Cancel);
  end else if Acol = GSLIB then
  begin
    TraiteLibelle (Acol,ARow,Cancel);
  end else if Acol = GSQTE then
  begin
    TraiteQte (Acol,ARow,Cancel);
    CalculeLaLigne(TOBL,TTgsDest);
  end else if Acol = GSPU then
  begin
    TraitePu (Acol,ARow,Cancel);
    CalculeLaLigne(TOBL,TTgsDest);
  end else if Acol = GSMARCHE then
  begin
    TraiteMarche (Acol,ARow,Cancel);
  end else if Acol = GSMONTANTD then
  begin
    TraiteMontantDest (Acol,ARow,Cancel);
    CalculeLaLigne(TOBL,TTgsDest);
  end;

  if (not Cancel) and (ZoneCalc(Acol)) then
  begin
    AjouteLaLigne (TOBL,TTgsDest,OldMt,'-');
    AjouteLaLigne (TOBL,TTgsDest,TOBL.GetDouble('MONTANT'));
    AfficheLaLigne(TOBL,TTgsDest ,Arow);
    StCellDest := GSDEST.Cells[Acol,Arow];
    DefiniPied;
  end else
  begin
    AffichelaLigne(TOBL,TTgsDest,ARow);
  end;
end;

procedure TOF_BTSAISTRFPOC.GSCellExitFrom(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var OldMt : double;
    TOBL : TOB;
begin
  if GSFROM.Cells[ACol,Arow] = StCellFrom then exit;
  TOBL := TOBFROM.detail[Arow-1];
  OldMt := TOBL.GetDouble('MONTANT');
  if Acol = GSQTE then
  begin
    TraiteQteFrom (Acol,ARow,Cancel);
  end else if Acol = GSMONTANTF then
  begin
    TraiteMontantFrom (Acol,ARow,Cancel);
  end;

  if (not Cancel) then
  begin
    AjouteLaLigne (TOBL,TTgsFrom,OldMt,'-');
    CalculeLaLigne(TOBL,TTgsFrom);
    AjouteLaLigne (TOBL,TTgsFrom,TOBL.GetDouble('MONTANT'));
  end;
  AffichelaLigne(TOBL,TTgsFrom,ARow);
  DefiniPied;
end;

procedure TOF_BTSAISTRFPOC.GSRowEnterDest(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  if Ou >= TOBDEST.Detail.count +1  then
  begin
    If (IsLigneComplete(Ou-1)) then AjouteDetail else Cancel := true;
  end;
end;

procedure TOF_BTSAISTRFPOC.GSRowExitDest(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BTSAISTRFPOC.BSELDestClick(Sender: TObject);
var Code : string;
begin
  Code := TOBTRFPOC.GetString ('BT2_PHASETRA');
  if SelectionPhase (TOBPiece.GetString('GP_AFFAIRE'),Code) then
  begin
  	TOBTRFPOC.SetString ('BT2_PHASETRA',Code);
    ChargeLesParags;
    DefiniLibelleDestination;
  end;

end;

function TOF_BTSAISTRFPOC.IsLigneComplete (Arow : Integer) : Boolean;
var TOBL : TOB;
begin
  Result := false;
  if Arow <= 1  then begin Result := True; Exit; end;
  if Arow > TOBDEST.detail.count +1 Then Exit;
  TOBL := TOBDEST.detail[Arow-1];
  if (TOBL.GetString('ARTSAISIE')<> '') and (TOBL.GetDouble('MONTANT')<>0) then Result := True;
end;

function TOF_BTSAISTRFPOC.AjouteDetail : TOB;
begin
  Result := TOB.Create ('BTRFDETAIL',TOBDEST,-1);
  AddChampsSupSaisie(Result);
end;


function TOF_BTSAISTRFPOC.IsExitInterdit(Acol,Arow : integer): boolean;
begin
  result := false;
  if (acol=0) or (Arow=0) then exit;
  if (Acol= 1) and (GSDEST.cells[Acol,Arow]='') and (GSDEST.row >= Arow) then
  begin
    result := true;
    exit;
  end;
  if (GSDEST.row > Arow) and (GSDEST.Cells[GSDEST.col,GSDEST.row]='') and (GSDEST.col<>1) then
  begin
    result := true;
  end;
end;


procedure TOF_BTSAISTRFPOC.TraiteCodeArt(ACol, Arow: Integer;var Cancel: boolean);
var RefArticle : string;
    found : Boolean;
    TOBA,TOBL : TOB;
begin
  cancel := false;
  TOBL := TOBDEST.Detail[GSDEST.row-1];
  TOBA := nil;
  RefArticle := GSDEST.Cells[ACol,Arow];
  if RefArticle <> '' then
  begin
    found := FindCodeArticle(RefArticle,TOBA);
    if found then
    begin
      SetArticleToLigne(TOBL,TOBA);
    end else
    begin
      if not RechercheArt (Arow) then
      begin
        GSDEST.cells[ACol,Arow] := '';
        cancel := True;
      end else
      begin
        CalculeLaLigne(TOBL,TTgsDest);
        AjouteLaLigne (TOBL,TTgsDest,TOBL.GetDouble('MONTANT'));
        AfficheLaLIgne(TOBL,TTgsDest ,GSDEST.row);
        StCellDest := GSDEST.Cells[GSDEST.col,GSDEST.row];
        DefiniPied;
      end;
    end;
  end;

end;

procedure TOF_BTSAISTRFPOC.TraiteLibelle(ACol, Arow: Integer;var Cancel: boolean);
var TOBL: TOB;
begin
  TOBL := TOBDEST.Detail[Arow-1];
  TOBL.SetString('LIBELLE',GSDEST.cells[Acol,Arow]);
end;

procedure TOF_BTSAISTRFPOC.TraiteMontantDest(ACol, Arow: Integer;var Cancel: boolean);
var TOBL: TOB;
begin
  TOBL := TOBDEST.Detail[Arow-1];
  TOBL.SetDOuble('MONTANT',Valeur(GSDEST.cells[Acol,Arow]));
  if TOBL.GetDouble('QTE') = 0 then TOBL.SetDouble('QTE',1);
  TOBL.SetDOuble('PA',ARRONDI(TOBL.GetDouble('MONTANT')/TOBL.GetDouble('QTE'),V_PGI.okdecP));
end;

procedure TOF_BTSAISTRFPOC.TraiteMontantFrom(ACol, Arow: Integer;var Cancel: boolean);
var TOBL: TOB;
    VAL : double;
begin
  TOBL := TOBFROM.Detail[Arow-1];
  //
  if Valeur(GSFROM.cells[Acol,Arow]) > TOBL.GetDouble('MONTANTINIT') then
  begin
    PGIInfo('Vous ne pouvez pas dépasser le montant initial');
    Cancel := True;
    Exit;
  end;
  TOBL.SetDouble('MONTANT',Valeur(GSFROM.cells[Acol,Arow]));
  if TOBL.GetDouble('PA') = 0 then TOBL.SetDouble('PA',1);
  VAL := ARRONDI(TOBL.GetDouble('MONTANT')/TOBL.GetDouble('PA'),V_PGI.okdecP);
  TOBL.SetDOuble('QTE',Val);
  //
  CalculeLaLigne(TOBL,TTgsFrom);
end;

procedure TOF_BTSAISTRFPOC.TraitePu(ACol, Arow: Integer;var Cancel: boolean);
var TOBL: TOB;
begin
  TOBL := TOBDEST.Detail[Arow-1];
  TOBL.SetDOuble('PA',Valeur(GSDEST.cells[Acol,Arow]));
end;

procedure TOF_BTSAISTRFPOC.TraiteQteFrom(ACol, Arow: Integer;var Cancel: boolean);
var TOBL: TOB;
begin
  TOBL := TOBFROM.Detail[Arow-1];
  if Valeur(GSFROM.cells[Acol,Arow]) > TOBL.getDouble('QTEINIT') then
  begin
    PgiInfo ('Vous ne pouvez pas dépassez la quantité initiale');
    Cancel := False;
    exit;
  end;
  TOBL.SetDOuble('QTE',Valeur(GSFROM.cells[Acol,Arow]));
end;

procedure TOF_BTSAISTRFPOC.TraiteQte(ACol, Arow: Integer;var Cancel: boolean);
var TOBL: TOB;
begin
  TOBL := TOBDEST.Detail[Arow-1];
  TOBL.SetDOuble('QTE',Valeur(GSDEST.cells[Acol,Arow]));
end;

procedure TOF_BTSAISTRFPOC.TraiteMarche(ACol, Arow: Integer;var Cancel: boolean);
begin
  GSDEST.Cells[Acol,Arow]:= StCellDest;
end;

procedure TOF_BTSAISTRFPOC.AjouteLaLigne(TOBL : TOB; TypeGrid: TTypeGridSaisie; Montant: double; Sens: string);
begin
  if TypeGrid = TTgsFrom then
  begin
    if Sens = '+' then
    begin
      MtInit := MtInit + Montant;
    end else
    begin
      MtInit := MtInit - Montant;
    end;
  end else
  begin
    if Sens = '+' then
    begin
      MtTransf := MtTransf + Montant;
    end else
    begin
      MtTransf := MtTransf - Montant;
    end;
  end;
end;

procedure TOF_BTSAISTRFPOC.CalculeLaLigne(TOBL: TOB; TypeGrid: TTypeGridSaisie);
begin
  if TypeGrid = TTgsFrom then
  begin
    TOBL.SetDouble('MONTANT',ARRONDI(TOBL.GetDouble('PA')*TOBL.GetDouble('QTE'),TFFacture(XX).DEV.decimale));
  end else
  begin
    TOBL.SetDouble('MONTANT',ARRONDI(TOBL.GetDouble('PA')*TOBL.GetDouble('QTE'),TFFacture(XX).DEV.decimale));
  end;
end;

procedure TOF_BTSAISTRFPOC.SetArticleToLigne(TOBX, TOBA: TOB);
var TOBL : TOB;
    MaxNumOrdre : integer;
begin
  TOBL := NewTOBLigne(TOBNewLigne,0);
  InitValoArtNomen (TOBA,'VEN');
  CalculePrixArticle (TOBA,'');
  InitLigneVide(TFFacture(XX).LaPieceCourante, TOBL, TFFacture(XX).TheTOBTiers, TFFacture(XX).TheTOBAffaire, 0, 1);
  TOBL.PutValue('GL_ENCONTREMARQUE', '-');
  if TOBPiece.getValue('GP_DOMAINE')<>'' then TOBL.PutValue('GL_DOMAINE', TOBPiece.GetValue('GP_DOMAINE'));
  MaxNumOrdre:=LireMaxNumOrdre(TFFacture(XX).LaPieceCourante);
  Inc(MaxNumOrdre);
  TOBL.PutValue('GL_NUMORDRE', MaxNumOrdre);
  EcrireMaxNumordre (TFFacture(XX).LaPieceCourante,MaxNumOrdre);
  PieceVersLigne(TFFacture(XX).LaPieceCourante, TOBL);
  AffaireVersLigne(TFFacture(XX).LaPieceCourante, TOBL, TFFacture(XX).TheTOBAffaire);
  //
  TOBA.PutValue('REFARTSAISIE', Trim(Copy(TOBA.GetValue('GA_ARTICLE'),1,18)));
  TOBL.PutValue('GL_ARTICLE', TOBA.GetValue('GA_ARTICLE'));
  TOBL.PutValue('GL_REFARTSAISIE', TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('GL_CODEARTICLE', TOBA.GetValue('GA_CODEARTICLE'));
  TOBL.PutValue('BNP_TYPERESSOURCE', TOBA.GetValue('BNP_TYPERESSOURCE'));
  TOBL.SetString('GL_DOMAINE',        TOBA.GetString('GA_DOMAINE'));
  //
  ArticleVersLigne (TFFacture(XX).LaPieceCourante,TOBA,TFFacture(XX).TheTOBConds,TOBL,TFFacture(XX).TheTOBTiers);
  TOBL.SetDouble('GL_QTEFACT',1);
  TOBL.SetDouble('GL_QTESTOCK',1);
  TOBL.SetDouble('GL_QTERESTE',1);
  TOBL.SetInteger('GL_NUMLIGNE',0);
  TOBX.Data := TOBL;
  AddChampsSupSaisie(TOBX);
end;


procedure TOF_BTSAISTRFPOC.ZoneSuivanteOuOk(GS : THgrid; var ACol, ARow: Longint; var Cancel: boolean);

  function ZoneAccessible(GS : THGrid; ACol, ARow: Longint): boolean;
  begin
    Result := GS.ColEditables [Acol];
  end;

var Sens, ii: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  cancel := false;
  OldEna := GS.SynEnabled;
  GS.SynEnabled := False;
  Sens := -1;
  ChgLig := (GS.Row <> ARow);
  ChgSens := False;
  if GS.Row > ARow then Sens := 1 else if ((GS.Row = ARow) and (ACol <= GS.Col)) then Sens := 1;
  ACol := GS.Col;
  ARow := GS.Row;
  ii := 0;
  while not ZoneAccessible(GS,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      if ChgLig then
      begin
        ACol := GS.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < GS.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := GS.FixedCols;
      end;
    end else
    begin
      if ((ACol = GS.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := GS.ColCount;
        ChgLig := False;
      end;
      if ACol > GS.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := GS.ColCount - 1;
      end;
    end;
  end;
  GS.SynEnabled := OldEna;
end;

procedure TOF_BTSAISTRFPOC.BeforePopup(Sender: Tobject);
begin
  if GSDEST.nbSelected > 0 then
  begin
    SetControlEnabled('MnAffectInterne',true);
    SetControlEnabled('MnAffectST',true);
  end else
  begin
    SetControlEnabled('MnAffectInterne',false);
    SetControlEnabled('MnAffectST',false);
  end;
end;

procedure TOF_BTSAISTRFPOC.AffecteST(Sender: TObject);
var TOBPARAM,TOBL : TOB;
    Fournisseur,ModePaie,CodeMarche : string;
    II : integer;
begin
  // Sélection du Sous traitant
  TOBParam := TOB.Create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('ACTION','SELECTION');
  TOBParam.AddChampSupValeur ('SOUSTRAIT','');
  TOBParam.AddChampSupValeur ('CODEMARCHE','');
  TOBParam.AddChampSupValeur ('AFFAIRE',TFFacture(XX).LaPieceCourante.GetValue('GP_AFFAIRE'));
  TOBParam.Data := TFFacture(XX).TheTOBSSTRAIT;
  TFFacture(XX).TheTOBSSTRAIT.Data := TFFacture(XX).ThePieceTrait;
  TheTOB := TOBParam;
  Fournisseur := '';
  ModePaie := '001';
  TRY
  	AGLLanceFiche('BTP','BTMARCHEST_MUL','BM1_AFFAIRE='+TFFacture(XX).LaPieceCourante.GetValue('GP_AFFAIRE'),'','');
    // --
    Fournisseur := TOBParam.GetValue('SOUSTRAIT');
    CodeMarche := TOBParam.GetValue('CODEMARCHE');
    if (Fournisseur <> '') and (CodeMarche <> '') then
    begin
      for II := 1 to GSDEST.Rowcount -1 do
      begin
        if GSDEST.IsSelected(II) then
        begin
          TOBL := TOBDEST.detail[II-1];
          TOBL.SetString('SOUSTRAITANT',Fournisseur);
          TOBL.SetString('CODEMARCHE',CodeMarche);
          AffichelaLigne (TOBL,TTgsDest,II);
        end;
      end;
    end;
    for II := 1 to GSDEST.Rowcount -1 do
    begin
      if GSDEST.IsSelected(II) then
      begin
        GSDEST.FlipSelection(II); 
      end;
    end;
  FINALLY
    TFFacture(XX).TheTOBSSTRAIT.Data := nil;
    TheTOB := nil;
  	TOBParam.free;
  end;
end;

procedure TOF_BTSAISTRFPOC.AffectInterneClick(sender: TObject);
begin

end;

procedure TOF_BTSAISTRFPOC.GSKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
var Vide: Boolean;
    Acol,Arow : Integer;
    cancel : boolean;
    TOBL : TOB;
begin
  Vide := (Shift = []);
  case Key of
    VK_RETURN: if (Vide) then
    begin
      Key := 0;
      SendMessage(GSDEST.Handle, WM_KeyDown, VK_TAB, 0);
    end;
    VK_DELETE: if (Shift = [ssCtrl]) then
    begin
      Key := 0;
      SupLigneClick;
    end;
    VK_F11 : if (Vide) then
    begin
      Key := 0;
      if GSDEST.Col = GSMONTANTD then
      begin
        TOBL := TOBDEST.detail[GSDEST.row-1];
        cancel := False;
        Acol := GSDEST.Col;
        Arow := GSDEST.Row;
        AjouteLaLigne (TOBL,TTgsDest,TOBL.GetDouble('MONTANT'),'-');
        DefiniPied;
        TOBL.SetDouble('MONTANT',0);
        TOBL.SetDouble('PA',0);
        GSDEST.Cells [GSDEST.Col,GSDEST.Row] := THNumEdit(GetControl('MTRESTE')).Text;
        GSCellExitDest(Self,Acol,Arow,cancel);
      end;
    end;
  end;
end;

procedure TOF_BTSAISTRFPOC.SupLigneClick;
var TOBL : TOB;
    Acol,Arow : Integer;
    Cancel : boolean;
begin
  if IsLigneComplete (GSDest.row) then
  begin
    TOBL := TOBDEST.Detail[GSDest.row-1];
    AjouteLaLigne (TOBL,TTgsDest,TOBL.GetDouble('MONTANT'),'-');
    DefiniPied;
    TOBL.Free;
    GSDEST.DeleteRow(GSDest.row);
    if GSDest.rowcount < 3  then
    begin
      TOBL := AjouteDetail;
      GSDEST.RowCount := TOBDEST.detail.count +2;
      TOBL.PutLigneGrid(GSDEST,GSDEST.RowCount -1,false,False,TheChampsDest);
    end;
    StCellDest := GSDEST.Cells[GSDEST.Col,GSDEST.Row];
    cancel := false;
    Arow := GSDEST.row;
    Acol := GSDEST.Col;
    GSRowEnterDest(Self,GSDEST.Row,cancel,false);
    GSCellEnterDest(Self,Acol,Arow,cancel);
  end;
end;

procedure TOF_BTSAISTRFPOC.RecalculeLesOuvrages(TOBRECALC : TOB);
var TOBL : TOB;
    II : Integer;
    TOBO : TOB;
begin
  for II := 0 TO TOBRECALC.detail.count - 1 do
  begin
    TOBL := TOB(TOBRECALC.Detail[II].Data);
    TOBO := TOBOUvrages.detail[TOBL.GetInteger('GL_INDICENOMEN')]; if TOBO = nil then Exit;
    RecalculeLigneFromOuvrage (TOBL,TOBO);
  end;
end;


procedure TOF_BTSAISTRFPOC.BeforeInsertTOB;
var II,NumLig : Integer;
    TOBT : TOB;
    TOBL,TT,TOBO,TOBLIGNE : TOB;
    GS : THGrid;
    TOBRECALC,TOBLR : TOB;
begin
  TOBRECALC := TOB.Create ('LES LIGNES A RECALC',nil,-1);
  GS := TFFacture(XX).GS;
  if ModeFonc = TmfModif then
  begin
    for II := TOBTRFPOC.Detail.count -1 downto 0 do
    begin
      TOBT := TOBTRFPOC.Detail[II];
      TOBL := TOB(TOBT.Data);
      if TOBL <> nil then
      begin
        if TOBT.GetString('BT3_TYPELIGNETRF')='001' then
        begin
          NumLig := TOBL.GetIndex + 1;
          if TOBL.GetString('GL_TYPELIGNE')='SD' then
          begin
            TOBLigne := FindLigneDocFromDetail (TOBPiece,TOBL);
            TOBLR := TOB.Create ('1 LIGNE',TOBRECALC,-1);
            TOBLR.AddChampSupValeur ('LIGNE',TOBLIGNE.GetInteger('GL_NUMLIGNE'));
            TOBLR.Data := TOBLIGNE;  
            if (TOBT.GetInteger('BT3_UNIQUEBLO')<>0) and (TOBT.GetString('BT3_CONTREP')='X') then
            begin
              if TOBL.GetInteger('GL_INDICENOMEN')<> 0 then
              begin
                TOBO := TOBOUvrages.detail[TOBL.GetInteger('GL_INDICENOMEN')-1];
                if TOBO <> nil then
                begin
                  TT := TOBOUvrages.FindFirst(['BLO_UNIQUEBLO'],[TOBT.GetInteger('BT3_UNIQUEBLO')],true);
                  if TT <> nil then TT.Free;
                end;
              end;
            end;
          end;
          GS.DeleteRow(NumLig);
          TOBL.Free;
        end else
        begin
          if TOBL.GetString('GL_TYPELIGNE')='SD' then
          begin
            TOBO := FindLigneOuvFromUnique (TOBL,TOBOUvrages,TOBL.GetInteger('UNIQUEBLO'));
            TOBO.SetDouble('MTTRANSFERT',arrondi(TOBO.GetDouble('MTTRANSFERT')- TOBT.GetDouble('BT3_MTTRANSFERT'),TFFacture(XX).DEV.decimale));
          end else
          begin
            TOBL.SetDouble('MTTRANSFERT',arrondi(TOBL.GetDouble('MTTRANSFERT')- TOBT.GetDouble('BT3_MTTRANSFERT'),TFFacture(XX).DEV.decimale));
          end;
        end;
      end;
    end;
  end;
  //
  RecalculeLesOuvrages(TOBRECALC);
  TOBTRFPOC.ClearDetail;
  TOBRECALC.Free;
end;

procedure TOF_BTSAISTRFPOC.AjouteLesDatas;

  procedure  AddSousTraitant(Fournisseur,CodeMarche : string);
  var TT : TOB;
  begin
    TT := TOB.Create('PIECEINTERV',TFFacture(XX).TheTOBSSTRAIT,-1);
		AddChampsSupTraitCreat (TT);
		AddChampsSupTrait (TT);
    TT.putValue('BPI_NATUREPIECEG',TOBpiece.GetString('GP_NATUREPIECEG'));
    TT.putValue('BPI_SOUCHE',TOBpiece.GetString('GP_SOUCHE'));
    TT.putValue('BPI_NUMERO',TOBpiece.GetInteger('GP_NUMERO') );
    TT.putValue('BPI_INDICEG',TOBpiece.GetInteger('GP_INDICEG'));
    TT.SetInteger('BPI_ORDRE',TFFacture(XX).TheTOBSSTRAIT.detail.count -1);
    //
    TT.putValue('BPI_TYPEINTERV','Y00');
    TT.putValue('BPI_TYPEPAIE','001');
    TT.putValue('BPI_DATECONTRAT',V_PGI.DateEntree);
    TT.SetBoolean('BPI_AUTOLIQUID',true);
    TT.SetString('BPI_FAMILLETAXE',GetInfoMarcheST(TOBpiece.GetString('GP_AFFAIRE'),Fournisseur,CodeMarche,'FAMILLETAXE1'));
    TT.SetString('BPI_TIERSFOU',Fournisseur);
    TT.SetString('BPI_CODEMARCHE',CodeMarche);
  end;

  procedure  AjouteSousTraitantPOC(TOBL: TOB; Fournisseur,CodeMarche : string);
  var TT : TOB;
  begin
    TT := TFFacture(XX).TheTOBSSTRAIT.FindFirst(['BPI_TYPEINTERV','BPI_TIERSFOU','BPI_CODEMARCHE'],['Y00',Fournisseur,CodeMarche],true);
    if TT = nil then
    begin
      AddSousTraitant(Fournisseur,CodeMarche);
    end;
    TOBL.putvalue('GL_FOURNISSEUR',Fournisseur);
    TOBL.putvalue('LIBELLEFOU',GetLibelleFou(Fournisseur));
    TOBL.putvalue('GLC_NATURETRAVAIL','002');
    TOBL.putvalue('GL_CODEMARCHE',CodeMarche);
  end;


  function FindLignePhaseFin : Integer;
  var TT : TOB;
      II : Integer;
      Val : Integer;
      NumOrdre : Integer;
  begin
    Result := -1;
    NumOrdre := -1;
    TT := TOBPhases.findfirst(['BPC_PHASETRA'],[TOBTRFPOC.GetString ('BT2_PHASETRA')],True);
    if TT  <> nil then
    begin
      for II := 9 downto 1 do
      begin
        Val := TT.GetInteger('BPC_PN'+InttoStr(II));
        if  Val <> 0 then
        begin
          NumOrdre := Val;
          break;
        end;
      end;
    end;
    //
    if NumOrdre < 0 then exit; 
    for II := 0 to TOBPiece.detail.count -1 do
    begin
      if TOBPiece.detail[II].GetInteger('GL_NUMORDRE')= NumOrdre then
      begin
        Result := RecupFinParagraph (TOBPiece,TOBPiece.detail[II].GetIndex,TOBPiece.detail[II].GetInteger('GL_NIVEAUIMBRIC'),true);
        break;
      end;
    end;
  end;

  function FindLigneSuivante (NumDepart : Integer) : Integer;
  var II : Integer;
      TOBL : TOB;
  begin
    Result := -1;
    TOBL := TFFacture (XX).LaPieceCourante.detail[NumDepart-1];
    if not IsOuvrage(TOBL) then BEGIN Result := NumDepart +1; Exit; END;
    for II := NumDepart+1 to TFFacture(XX).GS.rowCount -1 do
    begin
      TOBL := TFFacture (XX).LaPieceCourante.detail[II-1];
      if not IsOuvrage(TOBL) then BEGIN Result := II; break; Exit; END;
    end;
    if Result = -1 then Result := TFFacture(XX).GS.rowCount;
  end;

var II,LigD,LigS : Integer;
    TOBT,TOBTT,TOBL,TOBNL,TOBPAR,TOBO,TOBOS,TOBLigne : TOB;
    MaxNumOrdre : integer;
    CodeFamille2 : string;
    CoefPaPr,CoefPrPv,QteDejaTransf : double;
begin
  ConstituePhaseInterne;
  II := 0;
  // déduction ancien montant de la ligne départ pour le cumul
  // TRAITEMENT DES LIGNES PRISENT DANS LE TRANSFERT DEPART
  repeat
    TOBL := TOB(TOBFROM.detail[II].Data); if TOBL = nil then break;
    TOBT := TOBFROM.detail[II];
    TOBL.SetDouble('MTTRANSFERT',TOBL.GetDouble('MTTRANSFERT')+TOBT.GetDouble('MONTANT'));
    LigD := TOBL.GetIndex +1;
    //
    TOBT.SetDouble('BT3_MTTRANSFERT',TOBT.GetDouble('MONTANT'));
    TOBT.SetDouble('BT3_QTETRANSFERT',ARRONDI(TOBT.GetDouble('MONTANT')/TOBT.GetDouble('PA'),V_PGI.okdecQ));
    TOBT.SetInteger('BT3_UNIQUE',TOBTRFPOC.GetInteger('BT2_UNIQUE'));
    if TOBL.getString('GL_TYPELIGNE')='SD'  then
    begin
      LigS := LigD+1;
      TOBO := FindLigneOuvFromUnique (TOBL,TOBOUvrages,TOBL.GetInteger('UNIQUEBLO'));
      if TOBO <> nil then
      begin
        TOBLigne := FindLigneDocFromDetail (TOBPiece,TOBL);
        if TOBLigne <> nil then
        begin
          if TOBLigne.GetDouble('MTTRANSFERT') <> 0 then
          begin
            QteDejaTransf := Arrondi(TOBLigne.GetDouble('MTTRANSFERT')/TOBLigne.GetDouble('GL_DPA'),V_PGI.okdecQ);
          end;
          TOBLigne.SetDouble('MTTRANSFERT',TOBLigne.GetDouble('MTTRANSFERT')+TOBT.GetDouble('BT3_MTTRANSFERT'));
          TOBO.SetDouble('MTTRANSFERT',TOBT.GetDouble('BT3_MTTRANSFERT'));
          TOBT.SetInteger('BT3_UNIQUEBLO',TOBO.GetInteger('BLO_UNIQUEBLO'));
          //
          TOBT.SetBoolean('BT3_CONTREP',false);
          //
          TOBT.ChangeParent(TOBTRFPOC,-1);
          //
          TOBOS:=TOB.Create('LIGNEOUV',TobO.Parent,TOBO.GetIndex + 1);
          InsertionChampSupOuv (TOBOS,false);
          TOBTT := TOB.Create ('BTRFDETAIL',TOBTRFPOC,-1);
          //
          TOBOS.Dupliquer(TOBO,False,true);
          TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
          TOBOS.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
          TOBOS.SetInteger('NUMTRANSFERT',TOBTRFPOC.GetInteger('BT2_UNIQUE'));
          TOBOS.SetDouble('MTTRANSFERT',TOBT.GetDouble('BT3_MTTRANSFERT')*(-1));
          TOBOS.SetDouble('BLO_QTEFACT',ARRONDI(TOBT.GetDouble('BT3_QTETRANSFERT')/(TOBLigne.geTDouble('GL_QTEFACT')-QteDejaTransf),V_PGI.OkDecQ)*(-1));
          //
          TOBTT.setString('BT3_TYPELIGNETRF','001');
          TOBTT.SetDouble('BT3_MTTRANSFERT',TOBT.GetDouble('BT3_MTTRANSFERT'));
          TOBTT.SetDouble('MTTRANSFERT',TOBTT.GetDouble('BT3_MTTRANSFERT'));
          TOBTT.SetInteger('BT3_UNIQUE',TOBTRFPOC.GetInteger('BT2_UNIQUE'));
          TOBTT.SetInteger('BT3_UNIQUEBLO',TOBOS.GetInteger('BLO_UNIQUEBLO'));
          TOBTT.SetBoolean('BT3_CONTREP',true);
          //
          //
          CalculMontantHtDevLigOuv (TOBOS,TFFacture(XX).DEV);
        end;
        //
      end;
    end else
    begin
      LigS := FindLigneSuivante(LigD);
      TOBT.SetInteger('BT3_NUMORDRE',TOBL.GetInteger('GL_NUMORDRE'));
      TOBT.SetBoolean('BT3_CONTREP',false);
      //
      TOBT.ChangeParent(TOBTRFPOC,-1);
      TOBT.Data := TOBL;
      //
      TOBNL := NewTOBLigne(TOBPiece,LigS);
      TOBNL.Dupliquer(TOBL,true,true);
      MaxNumOrdre:=LireMaxNumOrdre(TOBPiece);
      Inc(MaxNumOrdre);
      TOBNL.PutValue('GL_NUMORDRE', MaxNumOrdre);
      EcrireMaxNumordre (TOBPiece,MaxNumOrdre);
      TOBTT := TOB.Create ('BTRFDETAIL',TOBTRFPOC,-1);
      TOBTT.setString('BT3_TYPELIGNETRF','001');
      TOBTT.SetDouble('BT3_MTTRANSFERT',TOBT.GetDouble('MONTANT')*(-1));
      TOBTT.SetInteger('BT3_NUMORDRE',TOBNL.GetInteger('GL_NUMORDRE'));
      TOBTT.SetInteger('BT3_UNIQUE',TOBTRFPOC.GetInteger('BT2_UNIQUE'));
      TOBTT.SetBoolean('BT3_CONTREP',true);
      TOBTT.data := TOBNL;
      //
      TOBNL.SetInteger('NUMTRANSFERT',TOBTRFPOC.GetInteger('BT2_UNIQUE'));
      TOBNL.SetDouble('MTTRANSFERT',TOBT.GetDouble('BT3_MTTRANSFERT')*(-1));
      TOBNL.SetDouble('GL_QTEFACT',TOBT.GetDouble('QTE')*(-1));
      TOBNL.SetDouble('GL_QTERESTE',TOBNL.GetDouble('GL_QTEFACT'));
      TOBNL.SetDouble('GL_QTESTOCK',TOBNL.GetDouble('GL_QTEFACT'));
      TOBNL.SetString('GL_RECALCULER','X');
      if IsOuvrage(TOBL) then
      begin
        DupliqueDetailOuv(TOBL,TOBNL);
      end;
      DelChampsSupSaisie(TOBT);
      DelChampsSupSaisie(TOBTT);
    end;
  until II >= TOBFROM.detail.count ;
  II := 0;
  //
  LigD := FindLignePhaseFin;
  if LigD <> -1 then
  begin
    TOBPAR := TOBPiece.detail[Ligd];
  end else
  begin
    Exit;
  end;
  repeat
    TOBL := TOB(TOBDEST.detail[II].Data); if TOBL = nil then break;
    TOBT := TOBDEST.detail[II];
    if (TOBT.getDouble('MONTANT')=0)  then BEGIN TOBT.free; continue; END;
    //
    TOBT.SetDouble('BT3_MTTRANSFERT',0);
    TOBT.SetInteger('BT3_UNIQUE',TOBTRFPOC.GetInteger('BT2_UNIQUE'));
    TOBT.setString('BT3_TYPELIGNETRF','001');
    TOBT.SetBoolean('BT3_CONTREP',false);
    //
    TOBT.ChangeParent(TOBTRFPOC,-1);
    //
    TOBNL := NewTOBLigne(TOBPiece,TOBPAR.getIndex+1);
    TOBNL.Dupliquer(TOBL,false,true);
    if TOBNL.Detail.count = 0 then
    begin
      NewTOBLigneFille(TOBL);
    end;
    MaxNumOrdre:=LireMaxNumOrdre(TOBPiece);
    Inc(MaxNumOrdre);
    EcrireMaxNumordre (TOBPiece,MaxNumOrdre);
    TOBNL.PutValue('GL_NUMORDRE', MaxNumOrdre);
    TOBT.SetInteger('BT3_NUMORDRE',MaxNumOrdre);
    TOBT.data := TOBNL;
    //
    if IsOuvrage(TOBNL) then
    begin
      AjouteDetailOuv(TOBNL);
      if TOBNL.GetDouble('GL_DPA')<> TOBT.GetDouble('PA') then
      begin
        TraitePrixOuvrage(TOBPiece,TOBNL,TFFacture(XX).TheTOBBases,TFFacture(XX).TheTOBBasesL,TFFacture(XX).TheTOBOuvrage, True, TOBT.GetDouble('PA'),0,TFFacture(XX).DEV,true);
      end;
    end;
    if (TOBT.GetString('SOUSTRAITANT') <> '') and (TOBT.GetString('CODEMARCHE') <> '') then
    begin
      AjouteSousTraitantPOC (TOBNL,TOBT.GetString('SOUSTRAITANT'),TOBT.GetString('CODEMARCHE'));
    end;
    //
    TOBNL.SetInteger('NUMTRANSFERT',TOBTRFPOC.GetInteger('BT2_UNIQUE'));
    TOBNL.SetInteger('GL_NIVEAUIMBRIC',TOBPAR.getInteger('GL_NIVEAUIMBRIC'));
    TOBNL.SetString('GL_LIBELLE',TOBT.GetString('LIBELLE'));
    TOBNL.SetDouble('GL_QTEFACT',TOBT.GetDouble('QTE'));
    TOBNL.SetDouble('GL_QTERESTE',TOBT.GetDouble('QTE'));
    TOBNL.SetDouble('GL_QTESTOCK',TOBT.GetDouble('QTE'));
    TOBNL.SetDouble('GL_DPA',TOBT.GetDouble('PA'));
    CodeFamille2 := GetInfoMarcheST (TOBNL.GetString('GL_AFFAIRE'),TOBT.GetString('SOUSTRAITANT'),TOBT.GetString('CODEMARCHE'),'FAMILLENIV2');
    if CodeFamille2 <> #0 then TOBNL.putvalue('GL_FAMILLENIV2',CodeFamille2);
    GetCoefPoc (TOBNL.GetString('GL_AFFAIRE'),CoefPaPr,CoefPrPv);
    if CoefPaPr = 0 then CoefPaPr := 1;
    if CoefPrPv = 0 then CoefPrPv := 1;
    if CoefPaPr <> 0 then
    begin
      TOBNL.putValue('GL_COEFFG',CoefPaPr-1);
      TOBNL.SetDouble('GL_DPR',ARRONDI(TOBNL.GetDouble('GL_DPA')*CoefPaPr,V_PGI.OkDecP));
    end;
    if CoefPrPv <> 0 then
    begin
      TOBNL.putValue('GL_COEFMARG',CoefPrPv);
      TOBNL.PutValue('POURCENTMARG',Arrondi((CoefPrPv-1)*100,2));
      TOBNL.SetDouble('GL_PUHT',ARRONDI(TOBNL.GetDouble('GL_DPR')*CoefPrPv,V_PGI.OkDecP));
      if TFFacture(XX).dev.code <> '' then
      begin
        TOBNL.SetDouble('GL_PUHTDEV',pivottodevise(TOBNL.GetDouble('GL_PUHT'),TFFacture(XX).DEV.Taux,TFFacture(XX).DEV.quotite,V_PGI.okdecP ));
      end else
      begin
        TOBNL.SetDouble('GL_PUHTDEV',TOBNL.GetDouble('GL_PUHT'));
      end;
    end;
    TOBNL.SetDouble('MTTRANSFERT',0);
    TOBNL.SetString('GL_RECALCULER','X');
    DelChampsSupSaisie(TOBT);
  until II >= TOBDEST.detail.count ;
  TOBPiece.SetString('GP_RECALCULER','X');
  TFFacture(XX).RecalculeDocumentClick(TFFacture(XX));
  TFFacture(XX).GS.Refresh;
end;

procedure TOF_BTSAISTRFPOC.AfterInsertTOB;
begin
  TFFacture(XX).SupLesLibDetail (TOBpiece);
end;

procedure TOF_BTSAISTRFPOC.DupliqueDetailOuv (TOBL,TOBNL: TOB);
var TOBO,TOBOS,TOBOD,TOBOuvrage,TOBOuv,TOBBases,TOBBasesL : TOB;
    II : Integer;
    NewPrix : Double;
    EnHT : boolean;
begin
  TOBOUvrage := TFFacture(XX).TheTOBOuvrage;
  TOBBases := TFFacture(XX).TheTOBBases;
  TOBBasesL := TFFacture(XX).TheTOBBasesL;
  EnHt := (TOBL.getValue('GL_FACTUREHT')='X');
  //
  TOBOUV := TOBOuvrage.detail[TOBL.GetInteger('GL_INDICENOMEN')-1];
  if TOBOUV <> nil then
  begin
    TOBO:=TOB.Create('',TOBOuvrage,-1) ;
    InsertionChampSupOuv (TOBO,false);
    TOBL.SetInteger('GL_INDICENOMEN',TOBO.GetIndex+1);
    //
    for II := 0 to TOBOUV.detail.count -1 do
    begin
      TOBOS := TOBOUV.detail[II];
      TOBOD := TOB.Create ('LIGNEOUV',TOBO,-1);
      TOBOD.Dupliquer(TOBOS,True,true);
      TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
      TOBOD.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
    end;
    NewPrix := TOBNL.GetDouble('GL_DPA');
    if TOBL.GetDouble('GL_DPA') <> TOBNL.GetDouble('GL_DPA') then
    begin
      TOBNL.SetDouble('GL_DPA',TOBL.GetDouble('GL_DPA'));
      TraitePrixOuvrage(TOBPiece,TOBNL,TOBBases,TOBBasesL,TOBOuvrage, EnHt, NewPrix,0,TFFacture(XX).DEV,True);
      ReinitCoefMarg (TOBNL,TOBO);
      if TOBNL.GetValue('GL_DPR')<>0 THEN
      begin
        TOBNL.PutValue('GL_COEFMARG',Arrondi(NewPrix/TOBNL.GetValue('GL_DPR'),4));
        TOBNL.PutValue('POURCENTMARG',Arrondi((TOBNL.GetValue('GL_COEFMARG')-1)*100,2));
      end else TOBNL.PutValue('GL_COEFMARG',0);

      if TOBNL.GetValue('GL_PUHT') <> 0 then
      begin
        TOBNL.PutValue('POURCENTMARQ',Arrondi(((TOBNL.GetValue('GL_PUHT')- TOBNL.GetValue('GL_DPR'))/TOBNL.GetValue('GL_PUHT'))*100,2));
      end else
      begin
        TOBNL.PutValue('POURCENTMARQ',0);
      end;
    end;
  end;
end;

procedure TOF_BTSAISTRFPOC.RecalculeLigneFromOuvrage (TOBL,TOBDEPART : TOB);
var   valeurs : T_Valeurs;
      EnHt : boolean;
begin
  EnHt := (TOBL.getValue('GL_FACTUREHT')='X');

  //
  InitTableau (Valeurs);
  CalculeOuvrageDoc (TOBDEPART,1,1,true,TFFacture(XX).DEV,valeurs,EnHt,true,true);
  StockeLesTypes(TOBL,Valeurs);
  if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
  begin
    TOBDEPART.putvalue ('BLO_DPA',valeurs[0]);
    TOBDEPART.putvalue ('BLO_DPR',valeurs[1]);
    TOBDEPART.PutValue('BLO_PMAP',valeurs[6]);
    TOBDEPART.PutValue('BLO_PMRP',valeurs[7]);
    TOBDEPART.putvalue ('BLO_PUHTDEV',valeurs[2]);
    TOBDEPART.putvalue ('BLO_PUTTCDEV',valeurs[3]);
    TOBDEPART.putvalue ('BLO_PUHT',devisetopivotEx(TOBDEPART.Getvalue ('BLO_PUHTDEV'),TFFacture(XX).DEV.Taux,TFFacture(XX).DEV.quotite,V_PGI.okdecP));
    TOBDEPART.putvalue ('BLO_PUTTC',devisetopivotEx(TOBDEPART.Getvalue ('BLO_PUTTCDEV'),TFFacture(XX).DEV.Taux,TFFacture(XX).DEV.quotite,V_PGI.OkdecP));
    TOBDEPART.putvalue ('BLO_PUHTBASE',TOBDEPART.GetValue('BLO_PUHT'));
    TOBDEPART.putvalue ('BLO_PUTTCBASE',TOBDEPART.getValue('BLO_PUTTC'));
    TOBDEPART.putvalue ('BLO_TPSUNITAIRE',valeurs[9]);
    TOBDEPART.Putvalue ('ANCPA',TOBDEPART.GetValue('BLO_DPA')) ;
    TOBDEPART.Putvalue ('ANCPR',TOBDEPART.GetValue('BLO_DPR')) ;
    //
    TOBDEPART.PutValue('GA_HEURE',valeurs[9]);
    TOBDEPART.PutValue('BLO_MONTANTPAFG',valeurs[10]);
    TOBDEPART.PutValue('BLO_MONTANTPAFR',valeurs[11]);
    TOBDEPART.PutValue('BLO_MONTANTPAFC',valeurs[12]);
    TOBDEPART.PutValue('BLO_MONTANTFG',valeurs[13]);
    TOBDEPART.PutValue('BLO_MONTANTFR',valeurs[14]);
    TOBDEPART.PutValue('BLO_MONTANTFC',valeurs[15]);
    TOBDEPART.PutValue('BLO_MONTANTPA',valeurs[16]);
    TOBDEPART.PutValue('BLO_MONTANTPR',valeurs[17]);
    //

    if EnHt then TOBDEPART.Putvalue('ANCPV',TOBDEPART.Getvalue('BLO_PUHTDEV'))
            else TOBDEPART.Putvalue('ANCPV',TOBDEPART.Getvalue('BLO_PUTTCDEV'));
    CalculMontantHtDevLigOuv (TOBDEPART,TFFacture(XX).DEV);
  end else
  begin
    GestionDetailPrixPose (TOBDepart);
  end;
  //
  TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
  TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
  TOBL.Putvalue('GL_PUHTNETDEV',valeurs[2]);
  TOBL.Putvalue('GL_PUTTCNETDEV',valeurs[3]);
  TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),TFFacture(XX).DEV.Taux,TFFacture(XX).DEV.quotite,V_PGI.OkdecP));
  TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),TFFacture(XX).DEV.taux,TFFacture(XX).DEV.quotite,V_PGI.okdecP));
  TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
  TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
  TOBL.Putvalue('GL_PUHTNET',TOBL.getValue('GL_PUHT'));
  TOBL.Putvalue('GL_PUTTCNET',TOBL.GetValue('GL_PUTTC'));
  TOBL.Putvalue('GL_DPA',valeurs[0]);
  TOBL.Putvalue('GL_DPR',valeurs[1]);
  TOBL.Putvalue('GL_PMAP',valeurs[6]);
  TOBL.Putvalue('GL_PMRP',valeurs[7]);
  TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
  //
  TOBL.PutValue('GL_MONTANTPAFG',valeurs[10]);
  TOBL.PutValue('GL_MONTANTPAFR',valeurs[11]);
  TOBL.PutValue('GL_MONTANTPAFC',valeurs[12]);
  TOBL.PutValue('GL_MONTANTFG',valeurs[13]);
  TOBL.PutValue('GL_MONTANTFR',valeurs[14]);
  TOBL.PutValue('GL_MONTANTFC',valeurs[15]);
  TOBL.PutValue('GL_MONTANTPA',valeurs[16]);
  TOBL.PutValue('GL_MONTANTPR',valeurs[17]);
end;

procedure TOF_BTSAISTRFPOC.AjouteDetailOuv(TOBL: TOB);
var TOBN,TOBLN,TOBOuvrage,TOBArticles,TOBDEPART : TOB;
    CodeNomen,RefUnique,Depot,CodeDevCli,TypeArt : string;
    NiveauDepart : Integer;
    valeurs : T_Valeurs;
    EnHt : boolean;
    Pu : double;
begin
  TOBOUvrage := TFFacture(XX).TheTOBOuvrage;
  TOBArticles := TFFacture(XX).TheTOBArticles;
  //
  TOBN:=TOB.Create('',nil,-1) ;
  TOBN.addchampsupValeur('NOMENCLATURE',CodeNomen,false);
  TRY
    Pu := TOBL.getDouble('GL_DPA');
    Depot := TOBL.GetString('GL_DEPOT');
    CodeNomen := TOBL.GetString('GL_CODEARTICLE');
    RefUnique := TOBL.GetString('GL_ARTICLE');
    CodeDevCli:= TOBL.GetValue ('GL_DEVISE');
    TypeArt:=TOBL.GetValue('GL_TYPEARTICLE') ;
    EnHt := (TOBL.getValue('GL_FACTUREHT')='X');

    NiveauDepart := 1;
    DerouleOuvrage(TOBN,CodeNomen,RefUnique,Depot,TOBArticles,1,TFFacture(XX).DEV,TFFacture(XX).IsVenteAchat) ;
    if TOBN.detail.count > 0 then
    begin
      TOBLN:=TOB.Create('',TOBOuvrage,-1) ;
      InsertionChampSupOuv (TOBLN,false);
      TOBL.SetInteger('GL_INDICENOMEN',TOBLN.GetIndex+1);
      InitTableau (Valeurs);
      RenseigneValoOuvrage(TOBN,CodeDevCli,1,1,TFFacture(XX).DEV,Valeurs) ;
      if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
      begin
        TOBDEPART := TOB.Create('LIGNEOUV',TOBLN,-1) ;
        InsertionChampSupOuv (TOBDEPART,false);
        inc (NiveauDepart);
      end else
      begin
        TOBDEPART := TOBLN;
      end;
      OuvVersLigOuv(TOBPiece,TOBL,TOBN,TOBDEPART,TOBArticles,NiveauDepart,0,TFFacture(XX).DEV,EnHt) ;
      NumeroteLigneOuv (TOBLN,TOBL,1,1,0,0,0);
      if TOBL.GetValue('GL_TYPEARTICLE')='ARP' then
      begin
        recupInfoLigneOuv (TOBDEPART,TOBDEPART.Detail[0]);
        TOBDEPART.PutValue('BLO_NATUREPIECEG',TOBDEPART.detail[0].GetValue('BLO_NATUREPIECEG'));
        TOBDEPART.PutValue('BLO_QTEFACT',1) ;
        TOBDEPART.PutValue('BLO_NIVEAU',1) ;
        TOBDEPART.PutValue('BLO_NUMORDRE',1) ;
        TOBDEPART.PutValue('BLO_QTEDUDETAIL',1) ;
        TOBDEPART.PutValue('BLO_PRIXPOURQTE',1) ;
        GetValoDetail (TOBDEPART);
      end;
      RecalculeLigneFromOuvrage (TOBL,TOBDEPART);
      //
    end;
  FINALLY
    TOBN.Free;
  END;
end;

procedure TOF_BTSAISTRFPOC.BdeleteClick(Sender: TObject);
begin
  if PGIAsk('Etes-vous sûr de supprimer ce transfert ?') <> mryes then Exit;
  BeforeInsertTOB;
  NumeroteLignesGC(TFFacture(XX).GS,TOBPiece,false,true);
  TOBTRFPOC.SetString('OKOK','D');
  TToolbarButton97(GetCOntrol('BFerme')).Click;
end;

procedure TOF_BTSAISTRFPOC.ConstituePhaseInterne;
var IndiceCurrent,NivCurrent,Indice : integer;
    CurrentPhase,TOBCompteur,TOBL : TOB;
    TypeLigne,TypeArticle : string;

  function GetPhase (TOBL : TOB; Niveau : integer) : integer;
  begin
    if Niveau = 1 then result := TOBL.GetValue('PN1') else
    if Niveau = 2 then result := TOBL.GetValue('PN2') else
    if Niveau = 3 then result := TOBL.GetValue('PN3') else
    if Niveau = 4 then result := TOBL.GetValue('PN4') else
    if Niveau = 5 then result := TOBL.GetValue('PN5') else
    if Niveau = 6 then result := TOBL.GetValue('PN6') else
    if Niveau = 7 then result := TOBL.GetValue('PN7') else
    if Niveau = 8 then result := TOBL.GetValue('PN8') else
    if Niveau = 9 then result := TOBL.GetValue('PN9') else
    result := 0;
  end;

  function EncodePhase: string;
  var P1,P2,P3,P4,P5,P6,P7,P8,P9 : integer;
  begin
    result := '';
    P1 := GetPhase(TOBCompteur,1); P2 := GetPhase(TOBCompteur,2); P3 := GetPhase(TOBCompteur,3);
    P4 := GetPhase(TOBCompteur,4); P5 := GetPhase(TOBCompteur,5); P6 := GetPhase(TOBCompteur,6);
    P7 := GetPhase(TOBCompteur,7); P8 := GetPhase(TOBCompteur,8); P9 := GetPhase(TOBCompteur,9);
    if P1 > 0 then result := IntTostr(P1);
    if P2 > 0 then result := result +'.'+IntToStr(P2);
    if P3 > 0 then result := result +'.'+inttoStr(p3);
    if P4 > 0 then result := result +'.'+IntTostr(P4);
    if P5 > 0 then result := result +'.'+IntToStr(P5);
    if P6 > 0 then result := result +'.'+inttoStr(p6);
    if P7 > 0 then result := result +'.'+IntTostr(P7);
    if P8 > 0 then result := result +'.'+IntToStr(P8);
    if P9 > 0 then result := result +'.'+inttoStr(p9);
  end;

  procedure DefiniCOmpteur (Niveau : integer);
  begin
    if Niveau = 1 then
    begin
      TOBCompteur.putValue ('PN1',TOBCompteur.GetValue('PN1')+1);
      TOBCompteur.putValue ('PN2',0);
      TOBCompteur.putValue ('PN3',0);
      TOBCompteur.putValue ('PN4',0);
      TOBCompteur.putValue ('PN5',0);
      TOBCompteur.putValue ('PN6',0);
      TOBCompteur.putValue ('PN7',0);
      TOBCompteur.putValue ('PN8',0);
      TOBCompteur.putValue ('PN9',0);
    end else
    if Niveau = 2 then
    begin
      TOBCompteur.putValue ('PN2',TOBCompteur.GetValue('PN2')+1);
      TOBCompteur.putValue ('PN3',0);
      TOBCompteur.putValue ('PN4',0);
      TOBCompteur.putValue ('PN5',0);
      TOBCompteur.putValue ('PN6',0);
      TOBCompteur.putValue ('PN7',0);
      TOBCompteur.putValue ('PN8',0);
      TOBCompteur.putValue ('PN9',0);
    end else
    if Niveau = 3 then
    begin
      TOBCompteur.putValue ('PN3',TOBCompteur.GetValue('PN3')+1);
      TOBCompteur.putValue ('PN4',0);
      TOBCompteur.putValue ('PN5',0);
      TOBCompteur.putValue ('PN6',0);
      TOBCompteur.putValue ('PN7',0);
      TOBCompteur.putValue ('PN8',0);
      TOBCompteur.putValue ('PN9',0);
    end else
    if Niveau = 4 then
    begin
      TOBCompteur.putValue ('PN4',TOBCompteur.GetValue('PN4')+1);
      TOBCompteur.putValue ('PN5',0);
      TOBCompteur.putValue ('PN6',0);
      TOBCompteur.putValue ('PN7',0);
      TOBCompteur.putValue ('PN8',0);
      TOBCompteur.putValue ('PN9',0);
    end else
    if Niveau = 5 then
    begin
      TOBCompteur.putValue ('PN5',TOBCompteur.GetValue('PN5')+1);
      TOBCompteur.putValue ('PN6',0);
      TOBCompteur.putValue ('PN7',0);
      TOBCompteur.putValue ('PN8',0);
      TOBCompteur.putValue ('PN9',0);
    end else
    if Niveau = 6 then
    begin
      TOBCompteur.putValue ('PN6',TOBCompteur.GetValue('PN6')+1);
      TOBCompteur.putValue ('PN7',0);
      TOBCompteur.putValue ('PN8',0);
      TOBCompteur.putValue ('PN9',0);
    end else
    if Niveau = 7 then
    begin
      TOBCompteur.putValue ('PN7',TOBCompteur.GetValue('PN7')+1);
      TOBCompteur.putValue ('PN8',0);
      TOBCompteur.putValue ('PN9',0);
    end else
    if Niveau = 8 then
    begin
      TOBCompteur.putValue ('PN8',TOBCompteur.GetValue('PN8')+1);
      TOBCompteur.putValue ('PN9',0);
    end else
    if Niveau = 9 then
    begin
      TOBCompteur.putValue ('PN9',TOBCompteur.GetValue('PN9')+1);
    end;
  end;


  procedure AjoutePhase (TOBL : TOB);
  var TOBphase,TOBLast : TOB;
  begin
    TOBLast := nil;
    if IndiceCurrent >= 0 then TOBLast := CurrentPhase;
    inc(NivCurrent);

    TOBphase := TOB.create ('PHASESCHANTIER',TOBPhases,-1);
    definiCompteur (nivcurrent);
    if TOBlast <> nil then
    begin
      TOBPhase.putValue('BPC_PN1',TobLast.getValue('BPC_PN1'));
      TOBPhase.putValue('BPC_PN2',TobLast.getValue('BPC_PN2'));
      TOBPhase.putValue('BPC_PN3',TobLast.getValue('BPC_PN3'));
      TOBPhase.putValue('BPC_PN4',TobLast.getValue('BPC_PN4'));
      TOBPhase.putValue('BPC_PN5',TobLast.getValue('BPC_PN5'));
      TOBPhase.putValue('BPC_PN6',TobLast.getValue('BPC_PN6'));
      TOBPhase.putValue('BPC_PN7',TobLast.getValue('BPC_PN7'));
      TOBPhase.putValue('BPC_PN8',TobLast.getValue('BPC_PN8'));
      TOBPhase.putValue('BPC_PN9',TobLast.getValue('BPC_PN9'));
      if NivCurrent = 9 then TOBPhase.putValue('BPC_PN9',TobL.getValue('GL_NUMORDRE')) else
      if NivCurrent = 8 then TOBPhase.putValue('BPC_PN8',TobL.getValue('GL_NUMORDRE')) else
      if NivCurrent = 7 then TOBPhase.putValue('BPC_PN7',TobL.getValue('GL_NUMORDRE')) else
      if NivCurrent = 6 then TOBPhase.putValue('BPC_PN6',TobL.getValue('GL_NUMORDRE')) else
      if NivCurrent = 5 then TOBPhase.putValue('BPC_PN5',TobL.getValue('GL_NUMORDRE')) else
      if NivCurrent = 4 then TOBPhase.putValue('BPC_PN4',TobL.getValue('GL_NUMORDRE')) else
      if NivCurrent = 3 then TOBPhase.putValue('BPC_PN3',TobL.getValue('GL_NUMORDRE')) else
      if NivCurrent = 2 then TOBPhase.putValue('BPC_PN2',TobL.getValue('GL_NUMORDRE'));
    end else
    begin
      TOBPhase.putValue('BPC_PN1',TobL.getValue('GL_NUMORDRE'));
    end;
    TOBPhase.putValue('BPC_AFFAIRE',TOBL.GetValue('GL_AFFAIRE'));
    TOBPhase.putValue('BPC_LIBELLE',TOBL.GetValue('GL_LIBELLE'));
    TOBPhase.putValue('BPC_PHASETRA',EncodePhase);
    indiceCurrent := TOBPhases.detail.count -1;
    CurrentPhase := TOBPhase;
  end;

  procedure RemontePhase;
  begin
    dec(IndiceCurrent);
    dec(NivCurrent);
    if (IndiceCurrent >= 0) and (NivCurrent > 0) then
    begin
      CurrentPhase := TOBPhases.detail[IndiceCurrent]
    end else
    begin
      CurrentPhase := nil;
      IndiceCurrent := -1;
      NivCurrent := 0;
    end;
  end;
  
  procedure AddSupPhases(TobPhase : TOB);
  begin
    TOBPhase.AddChampSupValeur('PN1',0);
    TOBPhase.AddChampSupValeur('PN2',0);
    TOBPhase.AddChampSupValeur('PN3',0);
    TOBPhase.AddChampSupValeur('PN4',0);
    TOBPhase.AddChampSupValeur('PN5',0);
    TOBPhase.AddChampSupValeur('PN6',0);
    TOBPhase.AddChampSupValeur('PN7',0);
    TOBPhase.AddChampSupValeur('PN8',0);
    TOBPhase.AddChampSupValeur('PN9',0);
  end;


begin
  TOBCompteur := TOB.create ('LE COMPTEUR DE PHASE',nil,-1);
  AddSupPhases (TOBCompteur);

  IndiceCurrent := -1;
  CurrentPhase := nil;

  TOBPhases.ClearDetail;
  for indice := 0 to TOBPIece.detail.count -1 do
  begin
    TOBL := TOBPIece.detail[Indice];
    if TOBL.getValue('GL_AFFAIRE')='' then continue;
    TypeLigne :=TOBL.GetValue('GL_TYPELIGNE');
    if copy(TypeLigne,1,2)='DP' then
    begin
      AjoutePhase (TOBL);
    end else if TypeLigne='ART' then
    begin
      TypeArticle := TOBL.GetString('GL_TYPEARTICLE');
      if (Copy(TypeArticle,1,2)='OU') and (TOBPiece.GetValue ('GP_NATUREPIECEG')='BCE') and (VH_GC.BTCODESPECIF = '001') then
      begin
        AjoutePhase (TOBL);
        RemontePhase;
      end;
    end else if copy(TypeLigne,1,2)='TP' then
    begin
      RemontePhase;
    end;
  end;
  
  TOBCompteur.Free;
end;

Initialization
  registerclasses ( [ TOF_BTSAISTRFPOC ] ) ;
end.
