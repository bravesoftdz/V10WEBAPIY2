{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 16/05/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : BTANALDEV ()
Mots clefs ... : TOF;BTANALDEV
*****************************************************************}
Unit UTOFBTANALDEV ;

Interface

Uses  StdCtrls,
      Controls,
      Classes,
      db,
      forms,
      sysutils,
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOF,
      stat,
      UTOB,
      AGLInit,
      SaisUtil,

      {$IFDEF EAGLCLIENT} maineagl,{$ELSE} Doc_Parser,DBCtrls, fe_main, {$ENDIF}
      NomenUtil,
      Factvariante,
      FactUtil,
      uBTPGetDocument,
      uEntCommun,
      FactComm,
      Paramsoc;

Type TSrcAnal = (SrcDoc, SrcPar, SrcOuv) ;
Type RAnalyseDoc = RECORD
                   CurLig,Debut,Fin:integer;
                   NivParagrapheMax : Integer;
                   paragraphe,Ouvrage,Libelle : string;
                   Cotraitance : Boolean;
                   END;
Type
  TOF_BTANALDEV = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
  end ;

var TOBResul      : TOB;
    MultiDoc      : Boolean;
    nTOBArticles  : Integer;
    nTOBOuvrages  : Integer;
    nTOBAffaires  : Integer;
    NumeroDevis   : String;
    CodeClient    : String;
    EtatDevis     : String;
    //
    CoefFG_Param  : Double;
    TauxHoraire   : Double;
    DEV           : RDEVISE;


procedure EntreeAnalyseDocument (Provenance: TsrcAnal;TOBArticles,TOBPieces,TOBOuvrage:TOB;arguments:array of variant;Nombre:integer);
procedure EntreeAnalyseOuvrageBib (Article:string;Nomenclature:string;Libelle:string); overload;
procedure EntreeAnalyseOuvrageBib (TobDet:Tob;Libelle:string); Overload;
procedure EntreeTestTob (TOBanalyse:TOB);
Procedure PrepaAnalyseMultiDoc (naturepiece, souche : string ; numero, indice : integer ; TobPieces : TOB);
Procedure AnalyseMultiDoc (TobPieces : TOB);

Implementation

uses  PiecesRecalculs,
      FactOuvrage;


procedure EntreeTestTob (TOBanalyse:TOB);
begin
TOBResul := TOBAnalyse;
AGLLanceFiche('BTP','BTANALDEV','','','LIBELLE= Analyse de tob memoire') ; // Lance TobViewer
TOBResul:=nil;
end;

procedure TOF_BTANALDEV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTANALDEV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTANALDEV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTANALDEV.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTANALDEV.OnArgument (S : String ) ;
var St1 : String;
begin
  Inherited ;
  st1 := Copy(S, 0, Pos('=',S) - 1);
  if st1 = 'LIBELLE' then
  begin
     st1 := Copy(S, Pos('=',S) + 1, 255);
     ecran.Caption := 'Analyse '+st1;
  end;

  if TOBresul <> nil then
  begin
     TFStat (Ecran).LaTOB := TOBResul;
  end;

end;

procedure TOF_BTANALDEV.OnClose ;
begin
  Inherited ;
end ;

// Géneralitée
procedure AjouteZoneTOB (var TOBTMP:TOB);
begin
TOBTMP.addchampsup('BAN_USER',false); TOBTMP.PutValue('BAN_USER','');
TOBTMP.addchampsup('BAN_NATURE',false);  TOBTMP.PutValue('BAN_NATURE','');
TOBTMP.addchampsup('BAN_LIBNATURE',false);TOBTMP.PutValue('BAN_LIBNATURE','');
TOBTMP.addchampsup('BAN_TYPE',false); TOBTMP.PutValue('BAN_TYPE','');
TOBTMP.addchampsup('BAN_LIBTYPE',false); TOBTMP.PutValue('BAN_LIBTYPE','');
TOBTMP.addchampsup('BAN_PAR',false); TOBTMP.PutValue('BAN_PAR','');
TOBTMP.addchampsup('BAN_SOUSPAR',false); TOBTMP.PutValue('BAN_SOUSPAR','');
TOBTMP.addchampsup('BAN_ARTICLE',false);  TOBTMP.PutValue('BAN_ARTICLE','');
TOBTMP.addchampsup('BAN_LIBELLE',false); TOBTMP.PutValue('BAN_LIBELLE','');
TOBTMP.addchampsup('BAN_QUALIFQTEVTE',false); TOBTMP.PutValue('BAN_QUALIFQTEVTE','');
TOBTMP.addchampsup('BAN_FAMILLENIV1',false); TOBTMP.PutValue('BAN_FAMILLENIV1','');
TOBTMP.addchampsup('BAN_FAMILLENIV2',false); TOBTMP.PutValue('BAN_FAMILLENIV2','');
TOBTMP.addchampsup('BAN_FAMILLENIV3',false); TOBTMP.PutValue('BAN_FAMILLENIV3','');
TOBTMP.addchampsup('BAN_QTEFACT',false); TOBTMP.PutValue('BAN_QTEFACT',0.0);
TOBTMP.addchampsup('BAN_VALACHAT',false);TOBTMP.PutValue('BAN_VALACHAT',0.0);
TOBTMP.addchampsup('BAN_COEFFG',false);  TOBTMP.PutValue('BAN_COEFFG',0.0);
TOBTMP.addchampsup('BAN_POUFG',false); TOBTMP.PutValue('BAN_POUFG',0.0);
TOBTMP.addchampsup('BAN_VALPR',false);  TOBTMP.PutValue('BAN_VALPR',0.0);
TOBTMP.addchampsup('BAN_COEFMARG',false); TOBTMP.PutValue('BAN_COEFMARG',0.0);
TOBTMP.addchampsup('BAN_POUMARG',false); TOBTMP.PutValue('BAN_POUMARG',0.0);
TOBTMP.addchampsup('BAN_VALPVHT',false); TOBTMP.PutValue('BAN_VALPVHT',0.0);
TOBTMP.addchampsup('BAN_VALPVTTC',false); TOBTMP.PutValue('BAN_VALPVTTC',0.0);
TOBTMP.addchampsup('BAN_VALMONTFG',false); TOBTMP.PutValue('BAN_VALMONTFG',0.0);
TOBTMP.addchampsup('BAN_VALMONTFR',false); TOBTMP.PutValue('BAN_VALMONTFR',0.0);
TOBTMP.addchampsup('BAN_VALMONTFC',false); TOBTMP.PutValue('BAN_VALMONTFC',0.0);
//
TOBTMP.addchampsup('BAN_FOURNPRINC',false); TOBTMP.PutValue('BAN_FOURNPRINC','');
//
end;

procedure CalculeCoef (var TOBResul: TOB);
Var Indice : Integer;
    Coef,Pourcent : Double;
    TOBLoc : TOB;
begin
for Indice := 0 To TobResul.detail.count - 1 do
begin
   TOBLoc := TOBResul.detail [Indice];
   if TOBLoc.getValue ('BAN_ARTICLE') = '' then TOBLOc.Free;
   if TOBLOC.getValue('BAN_VALACHAT') <> valeur ('0') then
   begin
      Coef := TOBLOc.getValue ('BAN_VALPR') / TOBLoc.getValue ('BAN_VALACHAT');
      TOBLOC.putvalue('BAN_COEFFG',Coef);
      Pourcent := (coef -1) * 100;
      TOBLOC.putvalue ('BAN_POUFG',Pourcent);
   end;
   if TOBLOC.getValue('BAN_VALPR') <> valeur ('0') then
   begin
      Coef := TOBLOc.getValue ('BAN_VALPVHT') / TOBLoc.getValue ('BAN_VALPR');
      TOBLOC.putvalue('BAN_COEFMARG',Coef);
   end;
   if GetParamSocSecur('SO_BTGESTIONMARQ', False) then
   begin
     if TOBLOC.getValue('BAN_VALPVHT') <> valeur ('0') then
     begin
        Pourcent := Arrondi(((TOBLOc.getValue ('BAN_VALPVHT') - TOBLoc.getValue ('BAN_VALPR')) / TOBLOC.getValue('BAN_VALPVHT'))*100,4) ;
        TOBLOC.putvalue ('BAN_POUMARG',Pourcent);
     end;
   end else
   begin
     if TOBLOC.getValue('BAN_VALPR') <> valeur ('0') then
     begin
        Pourcent := Arrondi(((TOBLOc.getValue ('BAN_VALPVHT') - TOBLoc.getValue ('BAN_VALPR')) / TOBLOC.getValue('BAN_VALPR'))*100,4) ;
        TOBLOC.putvalue ('BAN_POUMARG',Pourcent);
     end;

   end;
end;
end;

////////////////////////////////// Traitement Analyse sur Ouvrage
procedure SommeDetail(TOBNomen,TOBInfoArt : TOB; var TOBResul : TOB; QteDuDetail:double);
var Qte         : Double;
    PrixAchat   : Double;
    PrixRevient : Double;
    PrixVente   : Double;
    PrixPourQte : Double;
    //MTPAFG      : Double;
    //MTPAFR      : Double;
    //MTPAFC      : Double;
    //PAFG        : Double;
    //PAFR        : Double;
    //PAFC        : Double;
    TOuvrage    : T_Valeurs;
    //Coef        : Double;
begin
  PrixPourQte:= TOBInfoArt.getvalue('GA_PRIXPOURQTE');

  if PrixPourQte = valeur('0') then PrixPourQte := Valeur ('1');
  If TobInfoArt.GetValue ('GA_TYPEARTICLE') = 'POU' then

  begin
  // Dans le cas ou l'article est un Pourcentage , Il faut recalculer les éléments de l'ouvrage
     ValoriseOuvrage (TOBNomen.getValue('GNL_NOMENCLATURE'),nil,nil,TOuvrage,false);
     TobInfoArt.PutValue('GA_CALCPRIXPR','');
     TobInfoArt.PutValue('GA_CALCPRIXHT','');
     TobInfoArt.PutValue('GA_DPA',TOuvrage[0]);
     TobInfoArt.PutValue('GA_DPR',TOuvrage[1]);
  end;

  if QteDudetail = valeur('0') then QteDudetail := Valeur ('1');
  Qte := TOBNomen.getValue('GNL_QTE')/QteDuDetail;

  // modif LS 3/06/03
  //if TOBInfoArt.getValue('GA_TYPEARTICLE') = 'PRE' then
  TOBResul.putvalue ('BAN_QTEFACT',Qte);//TOBNomen.getValue('GNL_QTE')/QteDuDetail);
  //
  // Récupération du Prix d'achat
  if TOBInfoART.GetValue ('GA_CALCPRIXPR') = 'DPA' then
    PrixAchat   := (TobInfoArt.GetValue('GA_DPA') * Qte) / PrixPourQte
  else if TOBInfoART.GetValue ('GA_CALCPRIXPR') = 'PAA' then
    PrixAchat   := (TobInfoArt.GetValue('GA_PAHT') * Qte) / PrixPourQte
  else if TOBInfoART.getvalue('GA_CALCPRIXPR') = 'PMA' then
    PrixAchat   := (TobInfoArt.GetValue('GA_PMAP') * Qte) / PrixPourQte
  else if (TOBResul.GetString('BAN_TYPE') = 'SAL') then   //Modif FV : Dev. prioritaire DSL le 05/06/2012
    PrixAchat   := Arrondi((TauxHoraire * Qte) / PrixPourQte,V_PGI.OkDecV)
  else
    PrixAchat   := (TobInfoArt.GetValue('GA_DPA') * Qte) / PrixPourQte;

  // Récupération du Prix de revient
  if TOBInfoart.getValue ('GA_CALCPRIXHT') = 'DPR' then
    PrixRevient := (TobInfoArt.GetValue('GA_DPR') * Qte) / PrixPourQte
  else if TOBInfoart.getValue ('GA_CALCPRIXHT') = 'PMR' then
    PrixRevient := (TobInfoArt.GetValue('GA_PMRP') * Qte) / PrixPourQte
  else
    PrixRevient := (TobInfoArt.GetValue('GA_DPR') * Qte) / PrixPourQte;

  if TOBNomen.getValue('GNL_PRIXFIXE') <> valeur ('0') then
    PrixVente   := (TOBNomen.getvalue('GNL_PRIXFIXE') * Qte) / PrixPourQte
  else
     PrixVente  := (TOBInfoArt.getvalue('GA_PVHT') * Qte) / PrixPourQte;
  //
  TOBresul.putvalue ('BAN_VALACHAT',TOBresul.getvalue('BAN_VALACHAT') + Prixachat);
  TOBresul.putvalue ('BAN_VALPR',TOBresul.getvalue('BAN_VALPR') + PrixRevient);
  TOBresul.putvalue ('BAN_VALPVHT',TOBresul.getvalue('BAN_VALPVHT') + PrixVente);
  TOBresul.putvalue ('BAN_VALPVTTC',0);

end;



procedure APlatOuv ( TOBNomen : TOB ; var TOBResul : TOB; QteDuDetail : double);
var
   TOBTmp,TOBInfoART : TOB;
   Indice : Integer;
   QteDuDetailB : Double;
   QQ : Tquery;
   Req : String;
begin

  if TobNomen.detail.count <> 0 then
  begin
    // Il y a du sous détail donc on parcours le sous détail
    for Indice := 0 To TOBNomen.detail.count - 1 do
    begin
      if TobNomen.getvalue('QTEDUDETAIL') > 0 then
         QteduDetailB := QteDuDetail * TOBNomen.getValue('QTEDUDETAIL')
      else
         QteDuDetailB := QteDuDetail;
      TOBTmp := TobNomen.detail[Indice];
      TOBTmp.putvalue('GNL_QTE',TOBTmp.getValue('GNL_QTE') * TOBNomen.getValue('GNL_QTE'));
      AplatOuv (TOBTmp,TobResul,QteduDetailB);
    end;
  end
  else
  begin
    // Il n'y a pas de sous detail donc on met à jour tobResul
    // On récupere d'abord les infos articles
    Req := 'SELECT GA_TYPEARTICLE,GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,';
    Req := Req + 'GA_LIBELLE,GA_DPA,GA_PAHT,GA_DPR,GA_PVHT,GA_PRHT,GA_PMAP,GA_PMRP,GA_QUALIFUNITEVTE';
    Req := Req + ',GA_PRIXPOURQTE,GA_CALCPRIXHT,GA_CALCPRIXPR,GA_PRIXPOURQTE,GA_NATUREPRES,';
    Req := Req + 'BNP_LIBELLE,BNP_TYPERESSOURCE, GA_FOURNPRINC';
    Req := Req + ' FROM ARTICLE LEFT OUTER JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES WHERE GA_ARTICLE="';
    Req := Req + TOBNomen.getValue ('GNL_ARTICLE') + '"';
    QQ := OpenSQL (Req,true);
    if not QQ.eof then
    begin
      TOBInfoArt := TOB.create ('',nil,-1);
      TOBInfoArt.SelectDB ('',QQ);
      TOBInfoArt.addChampSup ('LIBTYPE',false);
      if TOBInfoArt.getValue('GA_NATUREPRES') <> '' then
      begin
         TOBinfoart.putvalue('LIBTYPE',rechdom('AFTTYPERESSOURCE',TOBInfoArt.getValue('BNP_TYPERESSOURCE'),false));
      end
      else
      begin
         TOBInfoArt.putvalue('GA_NATUREPRES',TOBInfoArt.getValue('GA_FAMILLENIV1'));
         if TOBInfoArt.getValue('GA_FAMILLENIV1') <> '' then
            TOBInfoArt.putvalue('BNP_LIBELLE',rechdom('GCFAMILLENIV1',TOBInfoArt.getValue('GA_FAMILLENIV1'),false))
         else
            TOBInfoArt.putvalue('BNP_LIBELLE','Non Renseigné');
         TOBInfoArt.putvalue('BNP_TYPERESSOURCE','MAR');
         TOBInfoArt.putvalue('LIBTYPE','Fourniture');
      end;
      TOBTmp := TOBResul.FindFirst (['BAN_USER','BAN_PAR','BAN_SOUSPAR','BAN_NATURE','BAN_TYPE','BAN_ARTICLE'],
                [V_PGI.User,'',TOBInfoArt.getValue('GA_NATUREPRES'),
                TOBInfoArt.getValue('BNP_TYPERESSOURCE'),TOBNomen.getValue('GNL_ARTICLE')],true);
      // Pas d'element trouvé . On en crée un.
      if TobTmp = nil then
      begin
         TOBTmp := TOB.create ('Fille',TOBResul,-1);
         AjouteZoneTOB (TOBTMP);
         TOBTMP.putvalue('BAN_USER',        V_PGI.user);
         TOBTMP.putvalue('BAN_NATURE',      TOBInfoArt.getValue('GA_NATUREPRES'));
         TOBTMP.putvalue('BAN_LIBNATURE',   TOBInfoArt.getValue('BNP_LIBELLE'));
         TOBTMP.putvalue('BAN_TYPE',        TOBInfoArt.getValue('BNP_TYPERESSOURCE'));
         TOBTMP.putvalue('BAN_LIBTYPE',     TobInfoArt.getValue('LIBTYPE'));
         TOBTMP.putvalue('BAN_FAMILLENIV1', TobInfoArt.getValue('GA_FAMILLENIV1'));
         TOBTMP.putvalue('BAN_FAMILLENIV2', TobInfoArt.getValue('GA_FAMILLENIV2'));
         TOBTMP.putvalue('BAN_FAMILLENIV3', TobInfoArt.getValue('GA_FAMILLENIV3'));
         TOBTMP.putvalue('BAN_ARTICLE',     TOBNomen.getValue ('GNL_ARTICLE'));
         TOBTMP.putvalue('BAN_LIBELLE',     TOBNomen.getValue ('GNL_LIBELLE'));
         TOBTMP.putvalue('BAN_QUALIFQTEVTE',TOBInfoArt.getValue ('GA_QUALIFUNITEVTE'));
         TOBTMP.PutValue('BAN_FOURNPRINC',  TOBInfoArt.getValue ('GA_FOURNPRINC'));
      end;
      SommeDetail(TOBNomen,TOBInfoArt,TOBTmp,QteDuDetail);
      TOBInfoArt.free;
    end;
    ferme(QQ);
  end;

end;

function MiseAPlatOuv (CodeNomen: String;TOBNomen : TOB; Var TOBResul: TOB) : Boolean;
var
Indice : Integer;
TobTemp : TOB;
begin
result := true;
//  CodeNomen obligatoire
if CodeNomen = '' then Exit;
//  TobNomen non renseignée, on la charge
if TobNomen = nil then ChargeNomen(CodeNomen, True, TobNomen);
// Parcours Pour remplissage
for Indice := 0 To TOBnomen.detail.Count -1 do
begin
    TobTemp := TobNomen.Detail[Indice];
    APlatOuv(TobTemp, TobResul,1);
end;
end;


procedure EntreeAnalyseOuvrageBib (Article:string;Nomenclature:string;Libelle:string);
var
TOBDet: TOB;
begin
TOBDet := TOB.Create('', nil, -1);
//  Remplissage de la tob ligne de l'ouvrage
chargeNomen (Nomenclature,true,TOBDet);
if TOBDet.Detail.Count > 0 then
   begin
   TOBResul := TOB.create ('Ma TOB',nil,-1);
   if MiseAPlatOuv (Nomenclature,TOBDet,TOBResul) then
      begin
      CalculeCoef (TOBResul);
      V_PGI.AutoSearch := true;
      AGLLanceFiche('BTP','BTANALDEV','','','LIBELLE=Ouvrage '+Libelle) ; // Lance TobViewer
      V_PGI.AutoSearch := false;
      end;
      TOBResul.free;
   end;
TOBDet.free;
end;

procedure EntreeAnalyseOuvrageBib (TobDet:Tob;Libelle:string);
begin
   if TOBDet.Detail.Count > 0 then
   begin
      TOBResul := TOB.create ('Ma TOB',nil,-1);
      if MiseAPlatOuv ('Local',TOBDet,TOBResul) then
      begin
         CalculeCoef (TOBResul);
         V_PGI.AutoSearch := true;
         AGLLanceFiche('BTP','BTANALDEV','','','LIBELLE=Ouvrage '+Libelle) ; // Lance TobViewer
         V_PGI.AutoSearch := false;
      end;
      TOBResul.Free ;
   end;
end;

// ----- FIn des traitements Ouvrages

// -- Debut des traitements documents
procedure SommeDetailDoc(TOBPiece,TOBInfoArt : TOB; var TOBResul : TOB);
var Qte         : Double;
    PrixAchat   : Double;
    PrixRevient : Double;
    PrixVente   : Double;
    PrixVenteTTC: Double;
    //MtPAFG      : Double;
    //MTPAFC      : Double;
    //MTPAFR      : Double;
    PAFG        : Double;
    PAFR        : Double;
    PAFC        : Double;
    PrixPourQte : Double;
    //Coef        : Double;
    //CoefMarge   : Double;
begin

  PrixPourQte:= TOBPiece.getvalue('GL_PRIXPOURQTE');

  If TobPiece.GetValue ('GL_TYPEARTICLE') = 'POU' then
  begin
     TobInfoArt.PutValue('GA_CALCPRIXPR','');
     TobInfoArt.PutValue('GA_CALCPRIXHT','');
  end;

  if PrixPourQte = valeur('0') then PrixPourQte := Valeur ('1');

  // Modif LS le 03/06/03
  //if (TOBPiece.getValue('GL_TYPEARTICLE') = 'PRE') (*and (TOBResul.getValue('BAN_TYPE')='SAL')*) then
  Qte := TOBPiece.getValue('GL_QTEFACT')/PrixPourQte;
  //
  //Modif FV : Dev. prioritaire DSL le 05/06/2012
  if (TOBResul.GetString('BAN_TYPE') = 'SAL') AND (CoefFG_Param <> 0) then
  begin
    PrixAchat   := Arrondi(TauxHoraire * Qte,V_PGI.OkDecV);
    PrixRevient := Arrondi(CalculMtFraisFromLigne(TobPiece.Parent, TobPiece, TauxHoraire, CoefFG_Param-1, PAFG, PAFR, PAFC, false),V_PGI.OkDecV);
    PrixRevient := Arrondi(PrixRevient * Qte,V_PGI.OkDecV);
  end
  else
  begin
    PrixAchat := Arrondi(TobPiece.GetValue('GL_DPA') * Qte,4);
    PrixRevient := Arrondi(TobPiece.GetValue('GL_DPR') * Qte,4);

    PAFG    := TOBPiece.getValue('GL_MONTANTFG');
    PAFR    := TOBPiece.getValue('GL_MONTANTFR');
    PAFC    := TOBPiece.getValue('GL_MONTANTFC');
  end;

  PrixVente := TobPiece.getvalue('GL_MONTANTHTDEV');
  PrixVenteTTC := TobPiece.getvalue('GL_MONTANTTTCDEV');
  //

  TOBResul.putvalue ('BAN_QTEFACT',qte);

  TOBResul.putvalue ('BAN_QTEFACT',TOBPIece.getValue('GL_QTEFACT')/PrixPourQte);
  TOBresul.putvalue ('BAN_VALACHAT',TOBresul.getvalue('BAN_VALACHAT') + Prixachat);
  //
  TOBresul.putvalue ('BAN_VALMONTFG',TOBresul.getvalue('BAN_VALMONTFG') + PAFG);
  TOBresul.putvalue ('BAN_VALMONTFR',TOBresul.getvalue('BAN_VALMONTFR') + PAFR);
  TOBresul.putvalue ('BAN_VALMONTFC',TOBresul.getvalue('BAN_VALMONTFC') + PAFC);
  //
  TOBresul.putvalue ('BAN_VALPR',TOBresul.getvalue('BAN_VALPR') + PrixRevient);
  TOBresul.putvalue ('BAN_VALPVHT',TOBresul.getvalue('BAN_VALPVHT') + PrixVente);
  TOBresul.putvalue ('BAN_VALPVTTC',TOBresul.getvalue('BAN_VALPVTTC') + PrixVenteTTC);

end;


procedure SommeDetailOuvDoc(TOBDetOuv,TOBInfoArt,TOBResul : TOB;QtePrec,QteDuDetailPrec: double);
var Qte         : Double;
    PrixAchat   : Double;
    PrixRevient : Double;
    PrixVente   : Double;
    PrixVenteTTC: Double;
    //MtpAFG      : Double;
    //MTPAFC      : Double;
    //MTPAFR      : Double;
    PAFG        : Double;
    PAFR        : Double;
    PAFC        : Double;
    PrixPourQte : Double;
    QteDuDetail : Double;
    Ratio       : double;
    MontantFg,MontantFC,MOntantFr : double;
begin

  QteDuDetail := TOBDetOuv.getvalue('BLO_QTEDUDETAIL'); If QteDuDetail = 0 then QteDuDetail := 1;
  PrixPourQte:= TOBDetOuv.getvalue('BLO_PRIXPOURQTE');

  If TobDetOuv.GetValue ('BLO_TYPEARTICLE') = 'POU' then
  begin
     TobInfoArt.PutValue('GA_CALCPRIXPR','');
     TobInfoArt.PutValue('GA_CALCPRIXHT','');
  end;

  if PrixPourQte = valeur('0') then PrixPourQte := Valeur ('1');

  Qte := (QtePrec * TOBDetOuv.getValue('BLO_QTEFACT')) /(PrixPourQte * QteDUdetail * QteDuDetailPrec);
  Ratio := (QtePrec) /(PrixPourQte * QteDUdetail * QteDuDetailPrec);

  //Modif FV : Dev. prioritaire DSL le 05/06/2012
  if (TOBDetOuv.GetString('BNP_TYPERESSOURCE') = 'SAL') AND (CoefFG_Param <> 0) then
  begin
    PrixAchat:= arrondi(TauxHoraire* Qte,4);
    PrixRevient := Arrondi(CalculePrixRevientOuv (TOBdetOuv,nil,TauxHoraire,CoefFG_Param, MontantFg,MontantFC,MontantFR) * Qte,4);

    PAFG    := Arrondi(MontantFG* ratio,4);
    PAFR    := Arrondi(MontantFr* ratio,4);
    PAFC    := Arrondi(MontanTfc* ratio,4);
  end else
  begin

    PrixAchat:= arrondi(TobDetOuv.GetValue('BLO_DPA') * Qte,4);
    PrixRevient := Arrondi(TobDetOuv.GetValue('BLO_DPR') * Qte,4);

    PAFG    := Arrondi(TobDetOuv.getValue('BLO_MONTANTFG')* ratio,4);
    PAFR    := Arrondi(TobDetOuv.getValue('BLO_MONTANTFR')* ratio,4);
    PAFC    := Arrondi(TobDetOuv.getValue('BLO_MONTANTFC')* ratio,4);
  //
  end;
  //
  PrixVente := Arrondi(TobDetOuv.getvalue('BLO_PUHTDEV')* qte,4);//V_PGI.okdecV);
  PrixVenteTTC := Arrondi(TobDetOuv.getvalue('BLO_PUTTCDEV')* qte,4);//V_PGI.OkdecV);
  //
  TOBResul.putvalue ('BAN_QTEFACT',Qte);

  TOBresul.putvalue ('BAN_VALACHAT',TOBresul.getvalue('BAN_VALACHAT') + Prixachat);
  TOBresul.putvalue ('BAN_VALPR',TOBresul.getvalue('BAN_VALPR') + PrixRevient);
  TOBresul.putvalue ('BAN_VALPVHT',TOBresul.getvalue('BAN_VALPVHT') + PrixVente);
  TOBresul.putvalue ('BAN_VALPVTTC',TOBresul.getvalue('BAN_VALPVTTC') + PrixVenteTTC);
  //
  TOBresul.putvalue ('BAN_VALMONTFG',TOBresul.getvalue('BAN_VALMONTFG') + PAFG);
  TOBresul.putvalue ('BAN_VALMONTFR',TOBresul.getvalue('BAN_VALMONTFR') + PAFR);
  TOBresul.putvalue ('BAN_VALMONTFC',TOBresul.getvalue('BAN_VALMONTFC') + PAFC);
  //
end;

Procedure RecupInfosDevis (TOBTMP, TobInitiale : TOB);
begin
  TOBTMP.AddChampsupValeur('BAN_NUMERODEVIS',NumeroDevis);
  TOBTMP.AddChampsupValeur('BAN_TIERS',CodeClient);
  TOBTMP.AddChampsupValeur('BAN_ETATDEVIS',EtatDevis);
end;

procedure TraiteLigneDoc (TOBARticles,TOBPiece : TOB; var TOBresul : TOB ;NomParagraphe,NomSousParagraphe: string);
var   TOBInfoArt,TOBTMP,TOBI : TOB;
      QQ              : TQuery;
      Req             : String;
      BTNatPrestation : string;
begin

  TOBInfoARt := nil;
  if TOBArticles <> nil then
  begin
    TOBI := TOBArticles.findFirst (['GA_ARTICLE'],[TOBPiece.getValue ('GL_ARTICLE')],True);
    if TOBI <> nil then
    begin
      TOBInfoArt := TOB.create ('',nil,-1);
      TOBINfoArt.Dupliquer (TOBI,false,true);
    end;
  end;

  BTNatPrestation := GetParamsocSecur('SO_BTNATPRESTATION', '');

  if TOBInfoArt = nil then
  begin
    Req := 'SELECT GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,';
    Req := Req + 'GA_CALCPRIXHT,GA_CALCPRIXPR,GA_PRIXPOURQTE,GA_NATUREPRES,';
    Req := req + 'GA_FOURNPRINC,';
    Req := Req + 'BNP_LIBELLE,BNP_TYPERESSOURCE FROM ARTICLE ';
    Req := Req + 'LEFT OUTER JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES WHERE GA_ARTICLE="';
    Req := Req + TOBPiece.getValue ('GL_ARTICLE') + '"';
    QQ := OpenSQL (Req,true);
    if not QQ.eof then
    begin
      TOBInfoArt := TOB.create ('',nil,-1);
      TOBInfoArt.SelectDB ('',QQ);
    end;
    ferme (QQ);
  end;

  if TOBPiece.getValue('GLC_NATURETRAVAIL') = '002' then
  begin
    TOBInfoArt.putvalue('GA_NATUREPRES',BTNatPrestation);
    TOBInfoArt.putvalue('BNP_LIBELLE', rechdom('BTNATPRESTATION',BTNatPrestation,false));
    //relecture prestation à partir de BTNatPrestation...
    TOBInfoArt.putvalue('BNP_TYPERESSOURCE','ST');
  end;

  if TOBInfoArt <> nil then
  begin
    TOBInfoArt.addChampSup ('LIBTYPE',false);

    if TOBInfoArt.getValue('GA_NATUREPRES') <> '' then
    begin
      TOBinfoart.putvalue('LIBTYPE',rechdom('AFTTYPERESSOURCE',TOBInfoArt.getValue('BNP_TYPERESSOURCE'),false));
    end else
    begin
      TOBInfoArt.putvalue('GA_NATUREPRES',TOBPiece.getValue('GL_FAMILLENIV1'));
      if TOBPiece.getValue('GL_FAMILLENIV1') <> '' then
      begin
        TOBInfoArt.putvalue('BNP_LIBELLE',rechdom('GCFAMILLENIV1',TOBPiece.getValue('GL_FAMILLENIV1'),false))
      end else
      begin
        TOBInfoArt.putvalue('BNP_LIBELLE','Non Renseigné');
      end;
      TOBInfoArt.putvalue('BNP_TYPERESSOURCE','MAR');
      TOBInfoArt.putvalue('LIBTYPE','Fourniture');
    end;

    // Recherche dans la tob si l'élément existe déjà.
    TOBTmp := TOBResul.FindFirst (['BAN_USER','BAN_PAR','BAN_SOUSPAR','BAN_NATURE','BAN_TYPE','BAN_ARTICLE'],
                                  [V_PGI.User,NomParagraphe,TOBInfoArt.getValue('GA_NATUREPRES'),
                                  TOBInfoArt.getValue('BNP_TYPERESSOURCE'),TOBPiece.getValue('GL_ARTICLE')],true);
    // Pas d'element trouvé . On en crée un.
    if TobTmp = nil then
    begin
      TOBTmp := TOB.create ('Fille',TobResul,-1);
      AjouteZoneTOB (TOBTMP);
      TOBTMP.putvalue('BAN_USER',         V_PGI.user);
      TOBTMP.putvalue('BAN_PAR',          NomParagraphe);
      TOBTMP.putvalue('BAN_SOUSPAR',      NomSousParagraphe);
      TOBTMP.putvalue('BAN_NATURE',       TOBInfoArt.getValue('GA_NATUREPRES'));
      TOBTMP.putvalue('BAN_LIBNATURE',    TOBInfoart.getValue('BNP_LIBELLE'));
      TOBTMP.putvalue('BAN_TYPE',         TOBInfoArt.getValue('BNP_TYPERESSOURCE'));
      TOBTMP.putvalue('BAN_LIBTYPE',      TobInfoArt.getValue('LIBTYPE'));
      TOBTMP.putvalue('BAN_FAMILLENIV1',  TobInfoArt.getValue('GA_FAMILLENIV1'));
      TOBTMP.putvalue('BAN_FAMILLENIV2',  TobInfoArt.getValue('GA_FAMILLENIV2'));
      TOBTMP.putvalue('BAN_FAMILLENIV3',  TobInfoArt.getValue('GA_FAMILLENIV3'));
      TOBTMP.putvalue('BAN_ARTICLE',      TOBPiece.getValue ('GL_ARTICLE'));
      TOBTMP.putvalue('BAN_LIBELLE',      TOBPiece.getValue ('GL_LIBELLE'));
      TOBTMP.putvalue('BAN_QUALIFQTEVTE', TOBPiece.getValue ('GL_QUALIFQTEVTE'));
      //FV1 : 21/03/2014-FS#915 - EN analyse devis, ajouter le fournisseur dans les champs disponibles
      if TOBPiece.GetString ('GL_FOURNISSEUR') <> '' then
        TOBTMP.putvalue('BAN_FOURNPRINC', TOBPiece.GetString('GL_FOURNISSEUR'))
      Else
        TOBTMP.putvalue('BAN_FOURNPRINC', TOBInfoArt.GetString('BAN_FOURNPRINC'));


      // Ajout pour analyse multi doc (BRL)
      if MultiDoc then RecupInfosDevis (TOBTMP, TOBPiece.Parent.Parent.Parent);
    end;
    SommeDetailDoc(TOBPiece,TOBInfoArt,TOBTmp);
    TOBInfoArt.free;
	end;
end;

procedure TraiteArticleOuv (TOBArticles,TOBDetOUV,TOBResul : TOB;NomParagraphe,NomSousParagraphe: string;QtePrec,QteDudetail : double; LibelleForce:string='');
var TOBInfoArt       : TOB;
    TOBTMP           : TOB;
    TOBI             : TOB;
    QQ               : TQuery;
    Req              : string;
    Article          : String;
    BTNatPrestation  : string;
begin

  TOBInfoArt := nil;

  Article := Trim( TOBDetOuv.getValue('BLO_ARTICLE'));

  if TOBArticles <> nil then
  begin
    TOBI := TOBArticles.findFirst(['GA_ARTICLE'],[TOBDetOUv.getValue ('BLO_ARTICLE')],true);
    if TOBI <> nil then
    begin
      TOBInfoArt := TOB.create ('',nil,-1);
      TOBINfoArt.Dupliquer (TOBI,false,true);
    end;
  end;

  BTNatPrestation := GetParamsocSecur('SO_BTNATPRESTATION', '');

  if TOBInfoArt = nil then
  begin
    Req := 'SELECT GA_FOURNPRINC, ';
    Req := Req + 'GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3,GA_CALCPRIXHT, ';
    Req := Req + 'GA_CALCPRIXPR,GA_PRIXPOURQTE,GA_NATUREPRES,BNP_LIBELLE,BNP_TYPERESSOURCE ';
    Req := Req + 'FROM ARTICLE ';
    Req := Req + 'LEFT OUTER JOIN NATUREPREST ON BNP_NATUREPRES=GA_NATUREPRES WHERE GA_ARTICLE="';
    Req := Req + TOBDetOUv.getValue ('BLO_ARTICLE') + '"';
    QQ := OpenSQL (Req,true);
    if not QQ.eof then
    begin
      TOBInfoArt := TOB.create ('',nil,-1);
      TOBInfoArt.SelectDB ('',QQ);
      ferme (QQ);
    end;
  end;

  if TobDetOuv.getValue('BLO_NATURETRAVAIL') = '002' then
  begin
    TOBInfoArt.putvalue('GA_NATUREPRES',BTNatPrestation);
    TOBInfoArt.putvalue('BNP_LIBELLE', rechdom('BTNATPRESTATION',BTNatPrestation,false));
    //relecture prestation à partir de BTNatPrestation...
    TOBInfoArt.putvalue('BNP_TYPERESSOURCE','ST');
  end;

  if TOBInfoArt <> nil then
  begin
    TOBInfoArt.addChampSup ('LIBTYPE',false);
    if TOBInfoArt.getValue('GA_NATUREPRES') <> '' then
    begin
      TOBinfoart.putvalue('LIBTYPE',rechdom('AFTTYPERESSOURCE',TOBInfoArt.getValue('BNP_TYPERESSOURCE'),false));
    end
    else
    begin
      TOBInfoArt.putvalue('GA_NATUREPRES',TOBDetOuv.getValue('BLO_FAMILLENIV1'));
      if TOBDetOUV.getValue('BLO_FAMILLENIV1') <> '' then
        TOBInfoArt.putvalue('BNP_LIBELLE',rechdom('GCFAMILLENIV1',TOBDetOuv.getValue('BLO_FAMILLENIV1'),false))
      else TOBInfoArt.putvalue('BNP_LIBELLE','Non Renseigné');
        TOBInfoArt.putvalue('BNP_TYPERESSOURCE','MAR');
      TOBInfoArt.putvalue('LIBTYPE','Fourniture');
    end;
    // Recherche dans la tob si l'élément existe déjà.
    TOBTmp := TOBResul.FindFirst (['BAN_USER','BAN_PAR','BAN_SOUSPAR','BAN_NATURE','BAN_TYPE','BAN_ARTICLE'],
                          [V_PGI.User,NomParagraphe,TOBInfoArt.getValue('GA_NATUREPRES'),
                          TOBInfoArt.getValue('BNP_TYPERESSOURCE'),Article],true);
    // Pas d'element trouvé . On en crée un.
    if TobTmp = nil then
    begin
      TOBTmp := TOB.create ('Fille',TobResul,-1);
      AjouteZoneTOB (TOBTMP);
      TOBTMP.putvalue('BAN_USER',       V_PGI.user);
      TOBTMP.putvalue('BAN_PAR',        NomParagraphe);
      TOBTMP.putvalue('BAN_SOUSPAR',    NomSousParagraphe);
      TOBTMP.putvalue('BAN_NATURE',     TOBInfoart.getValue('GA_NATUREPRES'));
      TOBTMP.putvalue('BAN_LIBNATURE',  TOBInfoart.getValue('BNP_LIBELLE'));
      TOBTMP.putvalue('BAN_TYPE',       TOBInfoArt.getValue('BNP_TYPERESSOURCE'));
      TOBTMP.putvalue('BAN_LIBTYPE',    TobInfoArt.getValue('LIBTYPE'));
      TOBTMP.putvalue('BAN_FAMILLENIV1',TobInfoArt.getValue('GA_FAMILLENIV1'));
      TOBTMP.putvalue('BAN_FAMILLENIV2',TobInfoArt.getValue('GA_FAMILLENIV2'));
      TOBTMP.putvalue('BAN_FAMILLENIV3',TobInfoArt.getValue('GA_FAMILLENIV3'));
      TOBTMP.putvalue('BAN_ARTICLE',    Article);

      TOBTMP.PutValue('BAN_FOURNPRINC', TOBDetOuv.getValue ('BLO_FOURNISSEUR'));
      if LibelleForce = '' then
      begin
        TOBTMP.putvalue('BAN_LIBELLE',TOBDetOuv.getValue ('BLO_LIBELLE'));
      end else
      begin
        TOBTMP.putvalue('BAN_LIBELLE',LibelleForce);
      end;
      TOBTMP.putvalue('BAN_QUALIFQTEVTE',TOBDetOuv.getValue ('BLO_QUALIFQTEVTE'));
      // Ajout pour analyse multi doc (BRL)
      if MultiDoc then RecupInfosDevis (TOBTMP, TOBDetOuv.Parent.Parent.Parent);
    end;
    SommeDetailOuvDoc(TOBDetOuv,TOBInfoArt,TOBTmp,QtePrec,QteDuDetail);
    TOBInfoArt.free;
  end;
end;

procedure TraiteDetOuv (TOBARticles,TOBOuv,TOBResul : TOB;NomParagraphe,NomSousParagraphe: string;QtePrec,QteDuDetailPrec:double; TheLibelle : string='');
var
   TobDetOuv: TOB;
   Indice : integer;
   Qte,QTeDuDetail : Double ;
   LeLibelle : string;
begin
for Indice := 0 to TOBOuv.detail.count - 1 do
    begin
    TOBDetOuv := TOBOuv.detail[Indice];
    if not IsVariante (TOBDetOuv) then
       begin
       if TOBDetOuv.detail.count > 0 then
          begin
          Qte := QtePrec* TOBDetOUV.GetValue('BLO_QTEFACT');
          QteDuDetail := TOBDetOuv.Getvalue('BLO_QTEDUDETAIL');
          if QteDuDetail = 0 then QteDuDetail := 1;
          QteDudetail := QteDuDetailPrec * QteDuDetail;
          if TheLibelle = '' then
          begin
          	if TOBDetOuv.getValue('BLO_TYPEARTICLE')='ARP' then LeLibelle := TOBDetOuv.getValue('BLO_LIBELLE');
          end else
          begin
          	leLIbelle := TheLibelle;
          end;
          TraiteDetOuv (TOBArticles,TOBDetOuv,TOBResul,NomParagraphe,NomSousParagraphe,Qte,QteDuDetail,LeLibelle);
          end else
          begin
          TraiteArticleOuv (TOBArticles,TOBDetOUV,TOBResul,NomParagraphe,NomSousParagraphe,QtePrec,QteDuDetailPrec,TheLibelle);
          end;
       end;
    end;
end;

procedure TraiteLigneOuv (TOBArticles,TOBPiece,TOBOuvrage : TOB; var TOBresul : TOB ;NomParagraphe,NomSousParagraphe: string);
var   TOBOuv: TOB;
      Qte,QTeDuDetail : Double;
      IndiceNomen : Integer;
      TheLibelle : string;
begin
  TheLibelle := '';

  if TOBOuvrage = nil then exit;
  IndiceNomen :=TOBPiece.getvalue('GL_INDICENOMEN');

  if IndiceNomen = 0 then exit;
  TOBOUv := TOBOuvrage.Detail[IndiceNomen -1];

  if TOBOUv <> nil then
  begin
    Qte := TOBPiece.getvalue('GL_QTEFACT');
    QteDudetail := 1;
    if TOBPIece.GetValue('GL_TYPEARTICLE')='ARP' then
    begin
      TheLibelle := TOBPIece.GetValue('GL_LIBELLE');
    end;
    TraiteDetOuv (TOBArticles,TOBOuv,TOBResul,NomParagraphe,NomSousParagraphe,Qte,QteDuDetail,TheLibelle);
  end;
end;

procedure AnalyseDocument (TOBArticles,TOBPieces,TOBOuvrage : TOB; var TOBresul: TOB; AnalyseDoc : RAnalyseDoc );
var Indice : Integer;
    ITypLig,ITypArt,IImbric, INatJob : Integer;
    NomParagraphe,NomSousParagraphe : String;
    TOBPiece: TOB;
    prevision : boolean;
    NumP1,Nump2,Nump3,Nump4,Nump5,Nump6,Nump7,Nump8,Nump9 : integer;

	procedure Init(niveau : integer);
  begin
  	if niveau = 1 then
    begin
    	NumP1:=0;
    	NumP2:=0;
    	NumP3:=0;
    	NumP4:=0;
    	NumP5:=0;
    	NumP6:=0;
    	NumP7:=0;
    	NumP8:=0;
    	NumP9:=0;
    end else if niveau = 2 then
    begin
    	Nump2:=0;
    	NumP3:=0;
    	NumP4:=0;
    	NumP5:=0;
    	NumP6:=0;
    	NumP7:=0;
    	NumP8:=0;
    	NumP9:=0;
    end;
  end;

  function ConstitueNumParagraphe (Nump1,Nump2,Nimp3,Nimp4,Nump5,Nump6,Nump7,Nump8,Nump9 : integer) : string;
  begin
  	Result := Format('%2d',[Nump1]);
  	if Nump2 > 0 then Result := result + '.'+Format('%2d',[Nump2]);
  end;

begin
// Init
	init (1);
//
  NomParagraphe := 'Courant';
  Prevision     := (TOBPieces.detail[0].GetValue ('GL_NATUREPIECEG')='PBT');
  ITypLig       := TOBPieces.detail[0].GetNumChamp ('GL_TYPELIGNE');
  ITypArt       := TOBPieces.detail[0].GetNumChamp ('GL_TYPEARTICLE');
  IImbric       := TOBPieces.detail[0].GetNumChamp ('GL_NIVEAUIMBRIC');
  INatJob       := TOBPieces.detail[0].GetNumChamp ('GLC_NATURETRAVAIL');

  //Modif FV : Dev. prioritaire DSL le 05/06/2012
  CoefFG_Param  := GetParamSocSecur('SO_COEFFG', 0);
  TauxHoraire   := GetParamSocSecur('SO_TAUXHORAIRE', 0);
  DEV.code      := Tobpieces.GetString('GP_DEVISE');
  GetInfosDevise(DEV);


  For Indice := AnalyseDoc.Debut to AnalyseDoc.Fin do
  begin
   TOBPiece := TOBPieces.detail[Indice];
   if (TOBPiece.getvaleur(ITypLig) <> 'ART') and
      ((not IsArticleVariante(TOBPiece))      or (not prevision)) and
   		(Not IsParagrapheStd(TOBPiece))        and
      ((not IsParagrapheVariante(TOBPiece))   or (not prevision)) then continue;

(* (copy(TOBPiece.Getvaleur(ITypLig),1,2) <> 'DP') and
   (copy(TOBPiece.Getvaleur(ITypLig),1,2) <> 'TP') then continue;
*)
   // Vérification Niveau si paragraphe
  //if (copy(TOBPiece.Getvaleur(ITypLig),1,2) = 'DP') and
   if (IsDebutParagraphe(TOBpiece)) and
      (TOBPiece.Getvaleur(IImbric) >= AnalyseDoc.NivParagrapheMax) then
      begin
        if TOBPiece.Getvaleur(IImbric) = 1 then
        begin
        	Inc(Nump1); init (2);
        	NomParagraphe := ConstitueNumParagraphe (Nump1,Nump2,Nump3,Nump4,Nump5,Nump6,Nump7,Nump8,Nump9)+': '+TOBPiece.getvalue('GL_LIBELLE')
        end else if TOBPiece.Getvaleur(IImbric) = 2 then
        begin
        	Inc(Nump2); init (3);
        	NomSousParagraphe := ConstitueNumParagraphe (Nump1,Nump2,Nump3,Nump4,Nump5,Nump6,Nump7,Nump8,Nump9)+':: '+TOBPiece.getvalue('GL_LIBELLE');
        end;
      (*
      // Sauvegarde du Nom du paragraphe
      if (TOBPiece.Getvaleur(IImbric) = AnalyseDoc.NivParagrapheMax) then
        NomParagraphe := ConstitueNumParagraphe (Nump1,Nump2,Nump3,Nump4,Nump5,Nump6,Nump7,Nump8,Nump9)+': '+TOBPiece.getvalue('GL_LIBELLE')
      if (TOBPiece.Getvaleur(IImbric) = AnalyseDoc.NivParagrapheMax+1) then
        NomSousParagraphe := ConstitueNumParagraphe (Nump1,Nump2,Nump3,Nump4,Nump5,Nump6,Nump7,Nump8,Nump9)+':: '+TOBPiece.getvalue('GL_LIBELLE');
      *)
      	continue;
      end;

//   if (copy(TOBPiece.Getvaleur(ITypLig),1,2) = 'TP') and
  if (IsFinParagraphe(TOBPiece)) and
     (TOBPiece.Getvaleur(IImbric) >= AnalyseDoc.NivParagrapheMax) then
  begin
    // Raz du Nom du paragraphe
    if (TOBPiece.Getvaleur(IImbric) = AnalyseDoc.NivParagrapheMax)   then NomParagraphe := 'Courant';
    if (TOBPiece.Getvaleur(IImbric) = AnalyseDoc.NivParagrapheMax+1) then NomSousParagraphe := 'Courant';
    continue;
  end;

//if (TOBPiece.Getvaleur(ITypLig) = 'ART') then
  if IsArticle(TOBPiece) then
  begin
    //si ligne cotraitant et pas de prise en compte de la cotraitance dans l'analyse...
    if ((TOBPiece.GetValeur (INatJob) = '001') AND (not AnalyseDoc.cotraitance)) then
      Continue
    else
    begin
      if ((TOBPiece.GetValeur (ITypArt)<> 'OUV') AND (TOBPiece.GetValeur (ITypArt) <> 'ARP'))  OR
         ((TOBPiece.GetValeur (ITypArt) = 'ARP') AND (TOBPiece.GetValue ('GL_INDICENOMEN')=0)) then
        TraiteLigneDoc (TOBArticles,TOBPiece,TOBResul,NomParagraphe,NomSousParagraphe)
      else
        TraiteLigneOuv (TOBArticles,TOBPiece,TOBOuvrage,TOBResul,NomParagraphe,NomSousParagraphe);
    end;
  end;

end;

end;

function OuvrageUniqueDocument (TOBPieces: TOB;Provenance: TSrcAnal; var Analyse: RanalyseDoc) : boolean;
var TOBLoc : TOB;
begin
result := false;
TOBLOc:= TOBPieces.FindFirst (['GL_ARTICLE'],[Analyse.ouvrage],true);
if TOBLOC <> nil then
   begin
   TOBLOc:= TOBPieces.findnext (['GL_ARTICLE'],[Analyse.ouvrage],true);
   if TOBLOc = nil then
      begin
      result := true;
      Analyse.Curlig := TOBLoc.GetIndex;
      end;
   end;
end;

function determineDebut ( TOBPiece: TOB; Provenance: TsrcAnal; var AnalyseDoc : RanalyseDoc; DepartRech : Integer) : boolean; overload;
var Indice : Integer;
    TOBLoc : TOB;
    Niveau : Integer;
begin
result := false;
if DepartRech <> -1 then
   begin
   if Provenance = SrcPar then
      begin
      TOBLoc := TOBPIece.detail[DepartRech];
      Niveau := TOBLOc.getvalue('GL_NIVEAUIMBRIC');
      for Indice := DepartRech downto 0 do
          begin
          // On recherche le 'DPx'
          TOBLoc := TOBPIece.detail[Indice];
//          if TOBLOC.getValue('GL_TYPELIGNE') = 'DP'+InttoStr (Niveau) then
						if IsDebutParagraphe (TOBLoc,Niveau) then
             begin
             AnalyseDoc.NivParagrapheMax := Niveau +1;
             AnalyseDoc.Debut := Indice;
             result := true;
             break;
             end;
          end;
      end else
      if Provenance = SrcOuv then
         begin
         AnalyseDoc.debut := DepartRech;
         result := true;
         end;
   end;
end;

function determineDebut ( TOBPiece: TOB; Provenance: TsrcAnal; var AnalyseDoc : RanalyseDoc; DepartRech : string) : boolean; overload;
var TOBLoc : TOB;
begin
result := false;
TOBLOC := TOBPiece.FindFirst (['GL_LIBELLE'],[DepartRech],true);
if TOBLOC <> nil then
   begin
   AnalyseDoc.debut := TOBLOC.GetIndex;
   AnalyseDoc.NivParagrapheMax := TOBLOc.getvalue('GL_NIVEAUIMBRIC')+1;
   result := true;
   end;
end;

function determinefin ( TOBPiece: TOB; Provenance: TsrcAnal; var AnalyseDoc : RanalyseDoc; DepartRech : Integer):boolean;
var TOBLoc : TOB;
    Indice : Integer;
    Niveau : Integer;
begin
result := false;
if DepartRech <> -1 then
   begin
   if Provenance = SrcPar then
      begin
      TOBLOC := TOBPiece.detail[DepartRech];
      Niveau := TOBLOc.getvalue('GL_NIVEAUIMBRIC');
      if Niveau <> 0 then
         begin
         for Indice := DepartRech to TOBPiece.detail.Count -1 do
             begin
             // On recherche le 'TPx'
             TOBLoc := TOBPIece.detail[Indice];
//             if TOBLOC.getValue('GL_TYPELIGNE') = 'TP'+InttoStr (Niveau) then
             if IsFinParagraphe(TOBLOC,Niveau) then
                begin
                AnalyseDoc.Fin := Indice;
                AnalyseDoc.NivParagrapheMax := TOBLOc.getvalue('GL_NIVEAUIMBRIC')+1;
                result := true;
                break;
                end;
             end;
         end;
      end else
      begin
      // SrcOuv
      AnalyseDoc.fin := DepartRech;
      result := true;
      end;
   end;
end;

function determineChampApplic ( Provenance: TsrcAnal;TOBPieces:TOB; var AnalyseDoc: RAnalyseDoc): boolean;
begin
result := false;
if (AnalyseDoc.debut <> -1) then
   begin
   if (AnalyseDoc.fin <> -1) then
      begin
      Result := true;
      end else
      begin
      if not (DetermineFin (TOBPieces,Provenance,AnalyseDoc,AnalyseDoc.debut)) then exit;
      Result := true;
      end;
   end else if analyseDoc.Fin <> -1 then
       begin
       if not (determineDebut (TOBPieces,Provenance,AnalyseDoc,AnalyseDoc.fin)) then exit;
       Result := true;
       end
   else if AnalyseDoc.CurLig <> -1 then
        begin
        if not (determineDebut(TOBPieces,Provenance,AnalyseDoc,AnalyseDoc.curlig)) then exit;
        if not (determineFin(TOBPieces,Provenance,AnalyseDoc,AnalyseDoc.Debut)) then exit;
        Result := true;
        end
   else if AnalyseDoc.paragraphe <> '' then
        begin
        if not (DetermineDebut(TOBPieces,Provenance,AnalyseDoc,AnalyseDoc.paragraphe)) then exit;
        if not (determineFin(TOBPieces,Provenance,AnalyseDoc,AnalyseDoc.Debut)) then exit;
        Result := true;
        end
   else if AnalyseDoc.Ouvrage <> '' then
        begin
        if OuvrageUniqueDocument (TOBPieces,Provenance,AnalyseDoc) then
           begin
           // Dans le cas ou l'ouvrage est référencé de maière unique
           result := true;
           end;
        end;
end;

procedure EntreeAnalyseDocument (Provenance: TsrcAnal;TOBArticles,TOBPieces,TOBOuvrage:TOB;arguments:array of variant;Nombre:integer);
var AnalyseDoc  : RAnalyseDoc;
    Indice      : Integer;
    chaine      : String;
    Valeur      : string;
begin

  AnalyseDoc.Debut := -1;
  AnalyseDoc.Fin := -1;
  AnalyseDoc.CurLig := -1;
  AnalyseDoc.NivParagrapheMax := -1;
  AnalyseDoc.paragraphe := '';
  AnalyseDoc.Ouvrage := '';
  AnalyseDoc.Libelle := '';
  AnalyseDoc.Cotraitance := false;

  MultiDoc := False;

  if TobPieces.detail.count = 0 then exit;

  for Indice := 0 to Nombre -1 do
  begin
    chaine := copy (arguments[Indice],0,pos('=',arguments[Indice])-1);
    Valeur := copy (arguments[Indice],pos('=',arguments[Indice])+1,255);
    if      chaine = 'LIBELLE'    then AnalyseDoc.Libelle     := Valeur
    else if chaine = 'CURLIG'     then AnalyseDoc.CurLig      := strtoInt (Valeur)
    else if chaine = 'DEBUT'      then AnalyseDoc.debut       := strtoInt (Valeur)
    else if chaine = 'FIN'        then AnalyseDoc.Fin         := strtoInt (Valeur)
    else if chaine = 'OUVRAGE'    then AnalyseDoc.ouvrage     := Valeur
    else if chaine = 'PARAGRAPHE' then AnalyseDoc.paragraphe  := Valeur
    else if chaine = 'COTRAITANT' then AnalyseDoc.Cotraitance := True;
  end;

  if Provenance = SrcDoc then
  begin
    AnalyseDoc.Debut  := 0;
    AnalyseDoc.fin := TOBPieces.detail.Count - 1;
    AnalyseDoc.NivParagrapheMax := 1;
  end
  else if Provenance = SrcPar then
  begin
    if (AnalyseDoc.paragraphe = '') and
    (AnalyseDoc.debut = -1) and
    (AnalyseDoc.fin = -1) and (AnalyseDoc.CurLig = -1) and
    (AnalyseDoc.NivParagrapheMax = -1) then exit;
    if not (determineChampApplic (Provenance,TOBPieces,AnalyseDoc)) then exit;
  end
  else if Provenance = SrcOuv then
  begin
    if not (determineChampApplic (Provenance,TOBPieces,AnalyseDoc)) then exit;
  end;

  // On peut maintenant lancer
  TOBResul := TOB.create ('MA Tob',nil,-1);

  AnalyseDocument (TOBArticles,TOBPieces,TOBOuvrage,TOBresul,AnalyseDoc);

  CalculeCoef (TOBResul);

  V_PGI.AutoSearch := true;

  AGLLanceFiche('BTP','BTANALDEV','','','LIBELLE= ' + AnalyseDoc.Libelle) ; // Lance TobViewer

  V_PGI.AutoSearch := false;

  TOBResul.free;

end;

Procedure PrepaAnalyseMultiDoc (naturepiece, souche : string ; numero, indice : integer ; TobPieces : TOB);
var
  TobPiece:TOB;
  Sql:string;
begin
// Remplissage de la TOB contenant les pieces selectionnées
TobPiece := Tob.Create ('PIECE', TobPieces,-1);
Sql := '"'+ NaturePiece
       + '";"'+ Souche
       + '";"'+ inttostr(Numero)
       + '";"'+ inttostr(Indice) + '"';

TobPiece.selectDB (sql, Nil, False);
end;

Procedure AnalyseMultiDoc (TobPieces : TOB);
var cledoc : string;
    TobP, TobParams : TOB;
    i, j : Integer;
    AnalyseDoc : RAnalyseDoc;
begin

  if TobPieces.detail.count = 0 then Exit;

  AnalyseDoc.CurLig := -1;
  AnalyseDoc.NivParagrapheMax := -1;
  AnalyseDoc.paragraphe := '';
  AnalyseDoc.Ouvrage := '';
  AnalyseDoc.Libelle := '';

  //Valeur forcé si l'on vient d'une sélection multi-devis... 
  AnalyseDoc.Cotraitance := True;

  MultiDoc := True;

  TobResul := TOB.Create ('MA TOB',nil,-1);

  for i:=0 to TobPieces.detail.count -1 do
  begin
    cledoc := EncodeRefPiece(TobPieces.detail[i]);

    TobP := TOB.Create ('UNE PIECE',nil,-1);
    TobParams := TOB.Create ('LES PARAMS',nil,-1);
    TobParams.AddChampSupValeur('CLEDOC',Cledoc);

    BTPGetDocument (TobParams, TobP);

    nTOBArticles := 0; nTOBOuvrages := 0; nTOBAffaires := 0;
    For j:=0 to TobP.Detail.count-1 do
    begin
      if TobP.Detail[j].NomTable = 'TOBARTICLES' then nTOBArticles := j
      else if TobP.Detail[j].NomTable = 'TOBOUVRAGES' then nTOBOuvrages := j
      else if TobP.Detail[j].NomTable = 'TOBAFFAIRES' then nTOBAffaires := j;
    end;

    AnalyseDoc.Debut := 0;
    AnalyseDoc.Fin := TobP.Detail[0].Detail[0].Detail.Count-1;
    if (nTOBArticles <> 0) and (nTOBOuvrages <> 0) then
    begin
      NumeroDevis:=   TobP.Detail[0].Detail[0].getValue('GP_NUMERO');
      CodeClient :=   TobP.Detail[0].Detail[0].getValue('GP_TIERS');
      if (nTOBAffaires <> 0) then EtatDevis := TobP.Detail[nTOBAffaires].Detail[0].getValue('AFF_ETATAFFAIRE');

      AnalyseDocument (TobP.Detail[nTOBArticles],TobP.Detail[0].Detail[0],TobP.Detail[nTOBOuvrages],TOBresul,AnalyseDoc);
    end;

    TobParams.Free;
    TobP.Free;
  end;
  CalculeCoef (TOBResul);
  V_PGI.AutoSearch := true;
  AGLLanceFiche('BTP','BTANALDEV','','','LIBELLE= '+AnalyseDoc.Libelle) ; // Lance TobViewer
  V_PGI.AutoSearch := false;
  TobResul.Free;
end;

Initialization
  registerclasses ( [ TOF_BTANALDEV ] ) ;
end.


