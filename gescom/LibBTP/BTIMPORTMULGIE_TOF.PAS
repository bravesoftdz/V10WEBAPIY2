{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/07/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTIMPORTMULGIE ()
Mots clefs ... : TOF;BTIMPORTMULGIE
*****************************************************************}
Unit BTIMPORTMULGIE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,fe_main,
{$else}
     eMul,
{$ENDIF}
     HTB97,
     paramSoc,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     Ent1,
     HMsgBox,
     vierge,
     Messages,
     Windows,
     SaisUtil,
     UTOF,
     UtilsParc,
     BTPUtil,
     AglInit ;

Type
  TOF_BTIMPORTMULGIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    fTobImport : TOB;
    fIdEntreprise : String;
    fEmplacIO : string;
    fListeChamps : string;
    flargeur : string;
    Falignement : string;
    FTitre : string;
    fNC : string;
    GS : Thgrid;
    fTOBFacture : TOB;
    fTOBresult : TOB;
    fTOBRapport : TOB;
    CodeEtab : string;
    fCodeArtDef,fArtDef : string;
    ferror : boolean;
    procedure AddlesChamps(OTOB : TOB);
    procedure AddlesChampsFacture (fTOBFacture : TOB);
    procedure AddlesSupligne (OTOB : TOB);
    procedure BSELECTALLClick (Sender : Tobject);
    procedure BVALIDEClick (Sender : Tobject);

    procedure ConstitueTOBImport;
    procedure DefiniAttributCol(GS: THgrid; Nom: string; Colonne: integer;LaLargeur, Lalignement, Letitre, LeNC: string);
    procedure GereAffichage;
    procedure Refresh;
    procedure TraiteFichier (FileName : string);
    procedure AfficheLaGrille (GS : Thgrid; TOBLignes : TOB; Depart : integer =0) ;
    procedure AfficheLaLigne (GS : THGrid; TOBLignes : TOB; Ligne : integer);
    function ControleSelected(OneTOB : TOB): boolean;
    function Generefacture(OneTOB: TOB): boolean;
    function GetCodetaux(RegimeTva : string; taux: double): string;
    procedure InitlesChampsFacture(fTOBFacture: TOB);
    function FindFourn(Emetteur: string; Auxi : boolean=true) : string;
    procedure ConstitueLesLignes(fTOBfacture,OneTOB : TOB);
    function GetRegimeTva(CodeFourn: string): string;
    function GetCodeArticle(fCodeArtDef: string): string;
    procedure ConstitueRapport(OKOK: boolean; TOBLue,TOBDepart, TOBresult: TOB);
    procedure MoveFilesOK(NomFile: string);
    procedure AddErreurRapport(TOBLUE: TOB;Themessage: string);

  end ;

Implementation

uses Grids,gerePiece;


procedure TOF_BTIMPORTMULGIE.AddlesChamps(OTOB: TOB);
begin
  OTOB.AddChampSupValeur('FILENAME','');
  OTOB.AddChampSupValeur('EMETTEUR','');
  OTOB.AddChampSupValeur('TIERS','');
  OTOB.AddChampSupValeur('AUXILIAIRE','');
  OTOB.AddChampSupValeur('LIBEMETTEUR','');
  OTOB.AddChampSupValeur('LIBINFO','');
  OTOB.AddChampSupValeur('NUMFACTURE','');
  OTOB.AddChampSupValeur('DATEFACTURE',iDate1900);
  OTOB.AddChampSupValeur('REFINTERNE','');
  OTOB.AddChampSupValeur('CHANTIER','');
  OTOB.AddChampSupValeur('BASE1',0);
  OTOB.AddChampSupValeur('BASE2',0);
  OTOB.AddChampSupValeur('BASE3',0);
  OTOB.AddChampSupValeur('BASE4',0);
  OTOB.AddChampSupValeur('BASE5',0);
  OTOB.AddChampSupValeur('TAUX1',0);
  OTOB.AddChampSupValeur('TAUX2',0);
  OTOB.AddChampSupValeur('TAUX3',0);
  OTOB.AddChampSupValeur('TAUX4',0);
  OTOB.AddChampSupValeur('TAUX5',0);
  OTOB.AddChampSupValeur('VALEUR1',0);
  OTOB.AddChampSupValeur('VALEUR2',0);
  OTOB.AddChampSupValeur('VALEUR3',0);
  OTOB.AddChampSupValeur('VALEUR4',0);
  OTOB.AddChampSupValeur('VALEUR5',0);
end;

procedure TOF_BTIMPORTMULGIE.AddlesChampsFacture (fTOBFacture : TOB);
begin
  fTOBFacture.AddChampSupValeur ('NATUREPIECEG', 'FF');
  fTOBFacture.AddChampSupValeur ('AFFAIRE', '');
  fTOBFacture.AddChampSupValeur ('TIERS', '');
  fTOBFacture.AddChampSupValeur ('ETABLISSEMENT', CodeEtab);
  fTOBFacture.AddChampSupValeur ('DOMAINE', '');
  fTOBFacture.AddChampSupValeur ('DATEPIECE', V_PGI.DateEntree);
  fTOBFacture.AddChampSupValeur ('REFINTERNE', '');
end;

procedure TOF_BTIMPORTMULGIE.InitlesChampsFacture (fTOBFacture : TOB);
begin
  fTOBFacture.SetString ('NATUREPIECEG', 'FF');
  fTOBFacture.SetString ('AFFAIRE', '');
  fTOBFacture.SetString ('TIERS', '');
  fTOBFacture.SetString ('ETABLISSEMENT', CodeEtab);
  fTOBFacture.SetString ('DOMAINE', '');
  fTOBFacture.SetDateTime ('DATEPIECE', V_PGI.DateEntree);
  fTOBFacture.SetString ('REFINTERNE', '');
end;

procedure TOF_BTIMPORTMULGIE.AddlesSupligne(OTOB: TOB);
begin
  OTOB.AddChampSupValeur ('TYPELIGNE', 'ART');
  OTOB.AddChampSupValeur ('ARTICLE', '');
  OTOB.AddChampSupValeur ('CODEARTICLE', '');
//  OTOB.AddChampSupValeur ('LIBELLE', '');
  OTOB.AddChampSupValeur ('QTEFACT', 1);
  OTOB.AddChampSupValeur ('AVECPRIX', 'X');
  OTOB.AddChampSupValeur ('PUHTDEV', 0);
end;

procedure TOF_BTIMPORTMULGIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULGIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULGIE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULGIE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULGIE.OnArgument (S : String ) ;
begin
  Inherited ;
  ferror := false;
  if VH^.EtablisDefaut <> '' then CodeEtab := VH^.EtablisDefaut;
  if VH^.EtablisCpta then
  begin
    if VH^.ProfilUserC[prEtablissement].Etablissement <> '' then CodeEtab := VH^.ProfilUserC[prEtablissement].Etablissement;
  end;
  //
  fTOBImport := TOB.Create ('LES FICHIERS',nil,-1);
  fTOBFacture := TOB.Create ('UNE FACTURE AC',nil,-1);
  fTOBresult := TOB.create ('LE RESULTAT',nil,-1);
  fTOBRapport := TOB.Create ('LE RAPPORT',nil,-1);
  AddlesChampsFacture (fTOBFacture);
  //

  GS := Thgrid(getControl('GS'));
  fListeChamps := 'INDICE;LIBEMETTEUR;LIBINFO;DATEFACTURE;REFINTERNE;CHANTIER;';
  flargeur := '2;60;60;8;50;10;';
  Falignement := 'G.0O X--;G.0O X--;G.0O X--;C.0  ---;G.0O X--;G.0O X--;';
  FTitre := ' ;Emetteur;Désignation;Date facture;Ref Interne;Chantier;';
  fNC := '0;0;0;0;0;0;'; //definition des zones saisissable ou non
  TToolbarButton97 (getControl('BSELECTALL')).OnClick := BSELECTALLClick;
  TToolbarButton97 (getControl('BVALID')).Onclick := BVALIDEClick;
  //
  fIdEntreprise := GetParamSocSecur ('SO_BTECHGEAN','');
  if fIdEntreprise = '' then
  begin
    PGIInfo ('ERREUR : Le Code EAN de votre société n''est pas défini !');
    ferror := true;
  end;
  fEmplacIO := GetParamSocSecur('SO_BTECHGEMPLAC','');
  if fEmplacIO = '' then
  begin
    PGIINfo ('ERREUR : Emplacement des échanges inter sociétés non défini !');
    ferror := true;
  end;
  fCodeArtDef := GetParamSocSecur('SO_BTECHGARTREC','');
  if fCodeArtDef = '' then
  begin
    PGIINfo ('ERREUR : l''article par défaut n''est pas renseigné !');
    ferror := true;
  end;
  fArtDef := GetCodeArticle(fCodeArtDef);
  if fArtDef = '' then
  begin
    PGIINfo ('ERREUR : l''article par défaut n''existe pas !');
    ferror := true;
  end;

  Refresh;

end ;

procedure TOF_BTIMPORTMULGIE.OnClose ;
begin
  fTOBImport.free;
  fTOBFacture.free;
  fTOBresult.free;
  fTOBRapport.free;
  Inherited ;
end ;

procedure TOF_BTIMPORTMULGIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULGIE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTIMPORTMULGIE.ConstitueTOBImport;
var FileName : string;
    Rec : TSearchRec;
    OneFacture : tob;
begin
  if (fIdEntreprise = '') or (fEmplacIO='') then exit;
  FileName:= fEmplacIO+'\*.ENV';
  if FindFirst (FileName,faAnyFile,Rec) = 0 then
  begin
    if (rec.name <> '.') and (rec.name <> '..') then
    begin
      TraiteFichier (fEmplacIO+'\'+Rec.Name);
    end;
    while FindNext (Rec) = 0 do
    begin
      if (rec.name <> '.') and (rec.name <> '..') then
      begin
        TraiteFichier (fEmplacIO+'\'+Rec.Name);
      end;
    end;
  end;
  SysUtils.FindClose (Rec);
  if fTOBIMport.detail.count = 0 then
  begin
    OneFacture := TOB.Create ('UNE FACTURE',fTobImport,-1);
    AddlesChamps(OneFacture);
  end;
end;


procedure TOF_BTIMPORTMULGIE.DefiniAttributCol(GS : THgrid; Nom: string; Colonne: integer; LaLargeur,Lalignement, Letitre, LeNC: string);
var Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec : integer;
    FF : string;
begin

  TransAlign(Lalignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
  GS.cells[Colonne,0] := leTitre;
  // Alignement dans la colonne
  if copy(LAlignement,1,1)='G' then GS.ColAligns[Colonne] := taLeftJustify
  else if copy(LAlignement,1,1)='D' then GS.ColAligns[Colonne] := taRightJustify
  else if copy(LAlignement,1,1)='C' then GS.ColAligns[Colonne] := taCenter;
  GS.ColWidths[Colonne] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
  GS.Collengths[Colonne] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
  if OkLib then
  begin
    GS.ColFormats[Colonne] := 'CB=' + Get_Join(Nom);
  end else if (Dec<>0) or (Sep) then
  begin
    GS.ColFormats[Colonne] := FF ;
  end;
  if LeNc = '0' then GS.ColEditables [Colonne]:= false;
end;


procedure TOF_BTIMPORTMULGIE.AfficheLaLigne (GS : THGrid; TOBLignes : TOB; Ligne : integer);
var TOBL : TOB;
begin
  TOBL := TOBlignes.detail[Ligne-1]; if TOBl = nil then exit;
  TOBL.PutLigneGrid (GS,Ligne,false,false,fListeChamps);
  GS.InvalidateRow(Ligne);
end;

procedure TOF_BTIMPORTMULGIE.AfficheLaGrille (GS : Thgrid; TOBLignes : TOB; Depart : integer =0) ;
var Indice : integer;
begin
  GS.RowCount := TOBLignes.detail.count+1; if GS.rowCount < 2 then GS.rowCount := 2;
  For Indice := Depart to TOBLignes.detail.count -1 do
  begin
    AfficheLaLigne (GS, TOBLignes, Indice+1);
  end;
  GS.Invalidate;
end;

procedure TOF_BTIMPORTMULGIE.GereAffichage;
var lesElements,lesLargeurs,lesAlignements,lesTitres,lesNC : string;
    Lelement,LaLargeur,Lalignement,LeTItre,LeNC : string;
    Colonne : integer;
begin
  GS.VidePile(false);
  GS.ColCount := 6; // Nombre de colonnes dans la saisie
  lesElements := fListeChamps;
  lesLargeurs := fLargeur;
  lesAlignements := Falignement;
  lesTitres := Ftitre;
  lesNc := fNC;
  //
  Lelement := ReadtokenSt(lesElements);
  LaLargeur := ReadtokenSt(lesLargeurs);
  Lalignement := ReadtokenSt(lesAlignements);
  LeTItre := readTokenSt(lesTitres);
  LeNC := readTokenSt(lesNc);
  Colonne := 0;

  fListeChamps := 'INDICE;LIBEMETTEUR;LIBINFO;DATEFACTURE;REFINTERNE;CHANTIER;';

  Repeat
    if Lelement = '' then break;
    DefiniAttributCol (GS,Lelement,Colonne,LaLargeur,Lalignement,Letitre,LeNC);
    //
    Lelement := ReadtokenSt(lesElements);
    LaLargeur := ReadtokenSt(lesLargeurs);
    Lalignement := ReadtokenSt(lesAlignements);
    LeTItre := readTokenSt(lesTitres);
    LeNC := readTokenSt(lesNc);
    inc(colonne);
    //
  until lelement = '';

  AfficheLaGrille(GS,fTobImport);
  TFVierge(ecran).HMTrad.ResizeGridColumns(GS);
end;

procedure TOF_BTIMPORTMULGIE.Refresh;
begin
  if ferror then exit;
  //
  fTOBImport.ClearDetail;
  TToolbarButton97 (getControl('BSELECTALL')).enabled := false;
  TToolbarButton97 (getControl('BVALID')).enabled := false;
  ConstitueTOBImport;
  GereAffichage;
  if fTobImport.detail.count > 0 then
  begin
    TToolbarButton97 (getControl('BSELECTALL')).enabled := true;
    TToolbarButton97 (getControl('BVALID')).enabled := true;
  end;
end;

procedure TOF_BTIMPORTMULGIE.TraiteFichier(FileName: string);
var FSource     : Textfile;
    MessageName,Emetteur,Destinataire,DatePiece,NumPiece,RefInterne,Chantier,Fourn : string;
    Base1,base2,base3,base4,base5 : double;
    taux1,taux2,taux3,taux4,taux5 : double;
    TVA1,TVA2,TVA3,TVA4,TVA5 : double;
    stInfos : string;
    StInfo,stContenu : string;
    iposEqual : integer;
    OneFacture : TOB;
begin
  AssignFile(FSource, FileName);
  Reset(FSource);
  try
    while not EOF(FSource) do
    begin
      Readln(FSource, StInfos);
      iPosEqual := Pos('=',StInfos);
      if iPosEqual > 0 then
      begin
        StInfo := copy(stInfos,1,iposEqual-1);
        stContenu := copy(stInfos,iposEqual+1,length(stInfos));
        if stInfo = 'MESSAGE' then MessageName := stContenu
        else if stInfo = 'EMETTEUR' then Emetteur := stContenu
        else if stInfo = 'DESTINATAIRE' then Destinataire := stContenu
        else if stInfo = 'NUMPIECE' then NumPiece := stContenu
        else if stInfo = 'DATE' then DatePiece := stContenu
        else if stInfo = 'REFINTERNE' then RefInterne := stContenu
        else if stInfo = 'CHANTIER' then Chantier := stContenu
        else if stInfo = 'BASE1' then base1 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'BASE2' then base2 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'BASE3' then base3 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'BASE4' then base4 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'BASE5' then base5 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'TAUX1' then taux1 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'TAUX2' then taux2 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'TAUX3' then taux3 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'TAUX4' then taux4 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'TAUX5' then taux5 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'VALEUR1' then TVA1 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'VALEUR2' then TVA2 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'VALEUR3' then TVA3 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'VALEUR4' then TVA4 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]))
        else if stInfo = 'VALEUR5' then TVA5 := StrToFloat(StringReplace (stContenu,'.',',',[rfreplaceall]));
      end;
    end;
  finally
    CloseFile(FSource);
  end;
  Fourn := FindFourn(Emetteur);
  if (MessageName <> 'INVOICE') or (Destinataire <> fIdEntreprise) or (Fourn = '') then exit;
  OneFacture := TOB.Create ('UNE FACTURE',fTobImport,-1);
  AddlesChamps(OneFacture);
  OneFacture.SetString('FILENAME',FileName);
  OneFacture.SetString('EMETTEUR',Emetteur);
  OneFacture.SetString('TIERS',FindFourn(Emetteur,false));
  OneFacture.SetString('AUXILIAIRE',Fourn);
  OneFacture.SetString('LIBEMETTEUR',rechdom('BTTIERSAUXI',Fourn,false,' AND T_NATUREAUXI="FOU"'));
  OneFacture.SetString('LIBINFO','Facture N° '+NumPiece+' du '+DatePiece);
  OneFacture.SetString('NUMFACTURE',NumPiece);
  OneFacture.SetDateTime('DATEFACTURE',StrToDateTime(DatePiece));
  OneFacture.SetString('REFINTERNE',RefInterne);
  OneFacture.SetString('CHANTIER',Chantier);
  //
  OneFacture.SetDouble('BASE1',base1);
  OneFacture.SetDouble('BASE2',base2);
  OneFacture.SetDouble('BASE3',base3);
  OneFacture.SetDouble('BASE4',base4);
  OneFacture.SetDouble('BASE5',base5);
  //
  OneFacture.SetDouble('TAUX1',taux1);
  OneFacture.SetDouble('TAUX2',taux2);
  OneFacture.SetDouble('TAUX3',taux3);
  OneFacture.SetDouble('TAUX4',taux4);
  OneFacture.SetDouble('TAUX5',taux5);
  //
  OneFacture.SetDouble('VALEUR1',TVA1);
  OneFacture.SetDouble('VALEUR2',TVA2);
  OneFacture.SetDouble('VALEUR3',TVA3);
  OneFacture.SetDouble('VALEUR4',TVA4);
  OneFacture.SetDouble('VALEUR5',TVA5);
end;

procedure TOF_BTIMPORTMULGIE.BSELECTALLClick(Sender: Tobject);
var II : integer;
begin
  for II := 1 to GS.RowCount do
  begin
    GS.FlipSelection(II); 
  end;
end;

procedure TOF_BTIMPORTMULGIE.BVALIDEClick(Sender: Tobject);
var II : integer;
    OneTOB : TOB;
    NbSelected : integer;
begin
  NBSelected := 0;
  For II := 1 to GS.rowCount do
  begin
    if GS.IsSelected(II) then
    begin
      if (II-1) <= (fTobImport.detail.count -1) then
      begin
        OneTOB := fTobImport.detail[II-1];
        inc(Nbselected);
        if ControleSelected (OneTOB) then
        begin
          Generefacture (OneTOB);
        end;
      end;
    end;
  end;
  if NbSelected <> 0 then
  begin
    TheTOB := fTOBRapport;
    AglLanceFiche ('BTP','BTRECALCPIECE_RAP','','','ERREURSOUPAS');
    TheTOB := Nil;
  end;
  Refresh;
end;

function TOF_BTIMPORTMULGIE.GetCodetaux (RegimeTva : string; taux: double): string;
var QQ : TQuery;
    Req, T1, T2, TvaLiquidSt : String;
    X : integer;
begin
  result := '';

  req := 'SELECT TV_CODETAUX FROM TXCPTTVA '+
         'WHERE '+
         'TV_TVAOUTPF="TX1" AND '+
         'TV_REGIME="'+regimeTVA+'" AND '+
         'TV_TAUXACH='+STRFpoint(Taux);

  TvaLiquidSt := Getparamsocsecur('SO_CODETVALIQUIDST','');
  if TvaLiquidSt <> '' then
    req := req + ' AND TV_CODETAUX NOT IN ('+ formatMultiValComboForSql(TvaLiquidSt) + ')';

  QQ := OpenSql (req,true,1,'',true);
  if not QQ.eof then result := QQ.fields[0].AsString;
  ferme (QQ);
end;

function TOF_BTIMPORTMULGIE.GetRegimeTva (CodeFourn : string) : string;
var QQ : TQuery;
begin
  result := '';
  QQ := OpenSql ('SELECT T_REGIMETVA FROM TIERS WHERE T_AUXILIAIRE="'+CodeFourn+'" AND T_NATUREAUXI="FOU"',true,1,'',true);
  if not QQ.eof then
  begin
    result:= QQ.fields[0].AsString;
  end;
  ferme (QQ);
end;

function TOF_BTIMPORTMULGIE.ControleSelected (OneTOB : TOB) : boolean;
var II : integer;
    RegimeTva,TheCode,Fourn : string;
begin
  result := false;
  Fourn := FindFourn(OneTOB.getString('EMETTEUR'));
  if fourn = '' then
  begin
    AddErreurRapport (OneTOB,'Fournisseur non défini');
    exit;
  end;
  if not ExisteSQL('SELECT 1 FROM AFFAIRE WHERE AFF_AFFAIRE="'+OneTOB.getString('CHANTIER')+'"') then
  begin
    AddErreurRapport (OneTOB,'Chantier non existant');
    exit;
  end;
  //
  RegimeTva := GetRegimeTva(Fourn);
  REsult := true;
  for II := 1 to 5 do
  begin
    if OneTOB.getDouble('BASE'+InttoStr(II))<> 0 then
    begin
      TheCode := GetCodetaux (regimeTva,OneTOB.GetDouble('TAUX'+InttoStr(II)));
      if TheCode = '' then
      BEGIN
        AddErreurRapport (OneTOB,'Anomalie intégration TVA');
        result := false;
        exit;
      END;
    end;
  end;
end;

procedure TOF_BTIMPORTMULGIE.AddErreurRapport ( TOBLUE : TOB; Themessage : string );
var TT : TOB;
begin
  TT := TOB.Create ('UNE LIGNE',fTOBRapport,-1);
  TT.AddChampSupValeur ('RAPPORT','Fichier '+TOBLue.getstring('FILENAME')+' : '+TheMessage + ' --> ERREUR');
end;

procedure TOF_BTIMPORTMULGIE.ConstitueRapport (OKOK : boolean; TOBLue,TOBDepart,TOBresult : TOB);
var debutTitre : string;
    TT : TOB;
begin
  TT := TOB.Create ('UNE LIGNE',fTOBRapport,-1);
  if OKOK then
  begin
    if TOBresult.detail.count > 1 then DebutTitre := 'Facture achat N° '+TOBresult.detail[1].GetString('GP_NUMERO')+' et Livraison N° '+TOBresult.detail[0].GetString('GP_NUMERO')
                                  else DebutTitre := 'Facture achat N° '+TOBresult.detail[0].GetString('GP_NUMERO');
    TT.AddChampSupValeur ('RAPPORT',DebutTitre + ' --> OK')
  end else
  begin
    DebutTitre := 'Intégration du fichier '+TOBLue.getstring('FILENAME');
    TT.AddChampSupValeur ('RAPPORT',DebutTitre + ' --> ERREUR');
  end;

end;

procedure TOF_BTIMPORTMULGIE.MoveFilesOK(NomFile : string);
var repert,OneFile,DestFile,SecFile,NomSec : string;
begin
  Repert := ExtractFilePath(NomFile);
  OneFile := ExtractFileName (NomFile);
  SecFile := copy (OneFile,1,Pos('.',OneFile)-1)+'.LET';
  NomSec := IncludeTrailingBackslash(IncludeTrailingBackslash(Repert))+SecFile;
  //
  DestFile := IncludeTrailingBackslash(IncludeTrailingBackslash(Repert)+'OK')+OneFile;
  if not DirectoryExists (IncludeTrailingBackslash(IncludeTrailingBackslash(Repert)+'OK')) then CreateDir (IncludeTrailingBackslash(IncludeTrailingBackslash(Repert)+'OK'));
  if FileExists(DestFile) then DeleteFile(PAnsiChar(DestFile));
  MoveFile(PansiChar(NomFile),PAnsiChar(DestFile));
  DestFile := IncludeTrailingBackslash(IncludeTrailingBackslash(Repert)+'OK')+SecFile;
  if FileExists(DestFile) then DeleteFile(PAnsiChar(DestFile));
  MoveFile(PansiChar(NomSec),PAnsiChar(DestFile));
end;

function TOF_BTIMPORTMULGIE.Generefacture (OneTOB : TOB) : boolean;
begin
  result := true;
  fTOBFacture.ClearDetail;
  fTOBresult.cleardetail;
  InitlesChampsFacture(fTOBFacture);
  fTOBFacture.setString('AFFAIRE',OneTOB.getString('CHANTIER'));
  fTOBFacture.setDateTime('DATEPIECE',OneTOB.getDatetime('DATEFACTURE'));
  fTOBFacture.SetString ('TIERS', OneTob.getString('TIERS'));
  fTOBFacture.SetString ('REFINTERNE', 'FF '+OneTOB.GetString('AUXILIAIRE')+'/'+OneTOB.getString('NUMFACTURE'));
  ConstitueLesLignes(fTOBfacture,OneTOB);
  if fTOBfacture.detail.count > 0 then
  begin
    Result := CreatePieceFromTOB (fTOBfacture,nil,nil,nil,'',fTOBresult);
    ConstitueRapport ((fTOBresult.detail.count <> 0),OneTOB,fTOBFacture,fTOBresult);
    if fTOBresult.detail.count > 0 then MoveFilesOK(OneTOB.GetString('FILENAME'));
  end;
end;

function TOF_BTIMPORTMULGIE.GetCodeArticle(fCodeArtDef : string) : string;
var Req : string;
    QQ : TQuery;
begin
  result := '';
  req := 'SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+fCodeArtDef+'"';
  QQ := OpenSql (req,true,1,'',true);
  if not QQ.eof then result := QQ.fields[0].AsString;
  ferme (QQ);
end;

function TOF_BTIMPORTMULGIE.FindFourn(Emetteur: string; Auxi : boolean=true): string;
var Sql : string;
    QQ : Tquery;
begin
  result:= '';
  Sql := 'SELECT T_TIERS,T_AUXILIAIRE FROM TIERS WHERE T_EAN="'+Emetteur+'" AND T_NATUREAUXI="FOU"';
  QQ := openSql (SQL,true,1,'',true);
  if not QQ.eof then
  begin
    if Auxi then result := QQ.fields[1].asstring
            else result := QQ.fields[0].asstring;
  end;
  ferme (QQ);
end;

procedure TOF_BTIMPORTMULGIE.ConstitueLesLignes(fTOBfacture, OneTOB: TOB);
var II : integer;
    regimeTva,TheCode : string;
    OneLigne : TOB;
begin
  RegimeTva := GetRegimeTva(OneTOB.getString('AUXILIAIRE'));
  for II := 1 to 5 do
  begin
    if OneTOB.getDouble('BASE'+InttoStr(II))<> 0 then
    begin
      TheCode := GetCodetaux (regimeTva,OneTOB.GetDouble('TAUX'+InttoStr(II)));
      if TheCode <> '' then
      BEGIN
        OneLigne := TOB.Create ('UNE LIGNE',fTOBfacture,-1);
        addlesSupligne (OneLigne);
        OneLigne.SetString('CODEARTICLE',fCodeArtDef);
        OneLigne.SetString('ARTICLE',fArtDef);
        OneLigne.SetDouble('PUHTDEV',StrToFloat(StringReplace (OneTOB.getString('BASE'+InttoStr(II)),'.',',',[rfreplaceall])));
        OneLigne.SetString('LIVDIRECTE','X');
      END;
    end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTIMPORTMULGIE ] ) ;
end.
