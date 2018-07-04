unit FactRg;

interface


Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
     HPdfPrev,UtileAGL,Maineagl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DB,
{$IFDEF V530}
      EdtEtat,EdtDoc,
{$ELSE}
      EdtREtat, EdtRDoc,
{$ENDIF}
{$ENDIF}
     SysUtils, Dialogs, Utiltarif, SaisUtil, UtilPGI, AGLInit, FactUtil,
     Math, StockUtil, EntGC, Classes, HMsgBox, FactNomen, FactTOB,
     Echeance, Paramsoc,factcomm,factcalc,FactPiece,FactRgpBesoin,
     uEntCommun,UtilTOBPiece
     ;

function RGMultiple (TOBpiece:TOB):boolean;
procedure InsereLigneRG(TOBPiece,TOBPieceRg,TOBBASEs,TOBBRG,TOBTIers,TOBAffaire,TOBLigneRG: TOB;NumIndice:integer);
procedure InsereLigneRGCot(TOBPiece,TOBPieceRG,TOBBASEs,TOBBRG,TOBTIers,TOBAffaire,TOBLigneRG: TOB;NumIndice:integer);
procedure RecalculeRG (TOBPorcs,TOBPIECE,TOBPIECERG,TOBBASES,TOBBASESRG,TOBPieceTrait : TOB;DEV:Rdevise);
function RecupDebutRG (TOBpiece : TOB;Indice : Integer): Integer;

procedure CalculBaseRG (TOBTiers,TOBPiece,TOBPieceTrait,TOBBasesL,TOBPIeceRG,TOBBasesRg,TOBS, TOBPorcs,TOBBASESG: TOB; R_Valeurs : R_CPercent;DEV:Rdevise;GenerationFac:boolean=false);
function GetCumultextRG (TOBPIeceRG:TOB;Nomchamp : string) : string;
function GetCumulValueRG (TOBPIeceRG:TOB;Nomchamp : string) : variant;
function GetCumulValueTaxeRG (TOBBasesRG:TOB;Nomchamp : string) : variant;
procedure GetCumulRG (TOBPieceRG : TOB; var SommePivot,SommeDevise: double;initial:boolean=false);
procedure FusionneTva (TOBTAXERG,TOBTAXEGLOB,TOBPIECERG,TOBTvaMillieme: TOB;Avoir : boolean=false);
procedure GetLesTaxesRG(TOBBasesRG :TOB; CatTaxe,FamTaxe:String; var RGD,RGP : double) ;
procedure GetLaTaxeRG(TOBB :TOB; var RGD,RGP : double) ;
procedure GetCumulTaxesRG(TOBBasesRG,TOBPieceRG :TOB; var RGD,RGP : double;DEV:Rdevise) ;
procedure GetMontantRG (TOBPIeceRG,TOBBasesRG : TOB;var XD,XP : double;DEV:Rdevise;HtForce:boolean=false;HorsCaution:boolean=false);
procedure ReIndiceLigneRg (TOBPieceRG: TOB;NumLigne,IndiceRg: integer);
procedure ReIndiceLigneBasesRg (TOBBasesRG: TOB;NumLigne,IndiceRg: Integer);
procedure initLigneRg (TOBRG,TOBPIECE : TOB; Interv : string='');
procedure ReajusteCautionDocOrigine (TOBPIece,TOBPieceRG:TOB; SensAjout : boolean;TOBPieceRG_O:TOB = nil);
procedure ReinitRGUtilises (TOBPieceRG : TOB);
function CalculPort (EnHt: boolean;TOBPorcs : TOB) : double;
function RestitueMontantCautionUtilise (TOBPieceRG : TOB) : boolean;
procedure EnregistreNumOrdreRG (TOBgenere,TOBPieceRg : TOB);
procedure GetCautionAlreadyUsed (TOBpiece,TOBL: TOB; PieceOrigine : string;var CautionUSed,CautionUsedDev : double; Intervenant : string ='';NewPiece : Boolean=false);
procedure GetReliquatCaution (TOBPieceRg : TOB; var RQ,RQDev : double) ;
procedure GetMontantRGReliquat(TOBPieceRg : TOB; var RGRP,RGRD : double; ForceTTC : boolean=false; UniquementsiTTC : boolean=false) ;
procedure SupprimeRg (TOBPieceRg,TOBBasesRg : TOB; IndiceRg : integer);
procedure VidePieceRGLigne (TOBPieceRG: TOB;Ligne : integer);
procedure RGVersLigne (TOBPiece,TOBPieceRG,TOBBasesRG,TOBLIG,TOBLigneRG : TOB; Applicretenue: boolean; GenerationFac : boolean=false);
function GetBaseCoTrait(TOBBases : TOB; Intervenant : string;EnHt : boolean) : double;
procedure CalculeRGSimple (TOBPorcs,TOBPIECE,TOBRG,TOBBASES,TOBBASESRG: TOB;DEV:Rdevise);
procedure CalculeRGSimpleCotrait (TOBPorcs,TOBPIECE,TOBPieceRG,TOBBASES,TOBBASESRG,TOBPieceTrait : TOB;DEV : RDevise);
function IsRGSimple (TOBPieceRG : TOB) : boolean;
procedure GetCautionUsedAfter (TOBPiece : TOB; PiecePrec : string;  NumSituation : integer;var CautionApres,CautionApresdev : double ;Intervenant : string = '');
procedure GetRg (TOBpieceRG : TOB; EnHt : boolean; Reliquat : boolean;  var Xp,Xd : double ; var NumCaution : string; Intervenant : string=''; PourTous : boolean=true);
procedure GetRGPRE (TOBRGPRE : TOB ; EnHt : boolean; var Xp,XD : double ; var numcaution : string);
procedure GetRGCUM (TOBRGPRE,TOBPieceRG : TOB ; EnHt,Reliquat : boolean; var Xp,XD : double ; var numcaution : string);
procedure RGMulVersLigne (TOBPiece,TOBPieceRG,TOBLIG,TOBLigneRG : TOB; Applicretenue: boolean; GenerationFac : boolean=false);
procedure ConstituePieceRg (TOBPiece,TOBpieceRg,TOBpieceTrait : TOB);
procedure GetTVARg (TOBpieceRG : TOB; EnHt : boolean; Reliquat : boolean;  var Xp,Xd : double ; Intervenant : string=''; PourTous : boolean=true;QuePourHT : boolean=false);
procedure ConstitueTVARG (TOBpieceRG,TOBBASESRG,TOBBRG : TOB);
procedure AnnuleCautionUtilise (TOBPiece,TOBPieceRG : TOB);
function FindRetenue (TOBPieceRG: TOB; NumOrdre : integer) : TOB;
function FindRGTOBLigne (TOBPiece : TOB ;NumOrdre : integer) : TOB;

const TRGModeInsertion : integer=0;
      TRGModeModif: integer=1;
type RGMode = (RGstd,RGDepass);
implementation
uses factvariante,factureBtp,FactLigneBase,UCotraitance,BTPUtil;

procedure GetRGPRE (TOBRGPRE : TOB ; EnHt : boolean; var Xp,XD : double ; var numcaution : string);
begin
	getRg (TOBRGPRE,EnHt,true,Xp,XD,numcaution);
end;

procedure GetRGCUM (TOBRGPRE,TOBPieceRG : TOB ; EnHt,Reliquat : boolean; var Xp,XD : double ; var numcaution : string);
var Indice : integer;
		TOBB  : TOB;
    MtRGResiduelle,ResiduDev,Residu : double;
begin
  NumCaution := '1111111';
  Xp := 0; XD := 0;
  residu := 0; ResiDuDev := 0;
  // Traitement sur situation précédente
  for Indice := 0 to TOBRGPRE.detail.count - 1 do
  begin
    TOBB := TOBRGPRE.detail[Indice];
    if EnHt then
    begin
      XD := XD + TOBB.GetDouble ('PRG_MTHTRGDEV');
      XP := XP + TOBB.GetDouble ('PRG_MTHTRG');
    end else
    begin
      if Reliquat then
      begin
        if (TOBB.GetString ('PRG_NUMCAUTION') <> '') then
        begin
          MtRGResiduelle := TOBB.GetDouble ('PRG_MTTTCRG') +
                            TOBB.GetDouble ('CAUTIONUTIL') -
                            TOBB.GetDouble ('CAUTIONAPRES') -
                            TOBB.GetDouble ('PRG_CAUTIONMT');
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
          XP := XP + MtRGResiduelle;
          //
          MtRGResiduelle := TOBB.GetDouble ('PRG_MTTTCRGDEV') +
                            TOBB.GetDouble ('CAUTIONUTILDEV') -
                            TOBB.GetDouble ('CAUTIONAPRESDEV') -
                            TOBB.GetDouble ('PRG_CAUTIONMTDEV');
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
          XD := XD + MtRGResiduelle;
        end else
        begin
          XP := XP + TOBB.GetDouble ('PRG_MTTTCRG');
          XD := XD + TOBB.GetDouble ('PRG_MTTTCRGDEV');
          if XD > 0 then NumCaution := '';
        end;
      end else
      begin
        XP := XP + TOBB.GetDouble ('PRG_MTTTCRG');
        XD := XD + TOBB.GetDouble ('PRG_MTTTCRGDEV');
        if (TOBB.GetDouble ('PRG_CAUTIONMT') > 0) then
        begin
          MtRGResiduelle := TOBB.GetDouble ('PRG_MTTTCRG') +
                            TOBB.GetDouble ('CAUTIONUTIL') -
                            TOBB.GetDouble ('CAUTIONAPRES') -
                            TOBB.GetDouble ('PRG_CAUTIONMT');
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
          Residu := Residu + MtRGResiduelle;
          //
          MtRGResiduelle := TOBB.GetDouble ('PRG_MTTTCRGDEV') +
                            TOBB.GetDouble ('CAUTIONUTILDEV') -
                            TOBB.GetDouble ('CAUTIONAPRESDEV') -
                            TOBB.GetDouble ('PRG_CAUTIONMTDEV');
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
          ResiduDev := ResiduDev + MtRGResiduelle;
        end else
        begin
          if XD > 0 then NumCaution := '';
        end;
      end;
    end;
  end;
  // traitement sur la situation courante
  for Indice := 0 to TOBPieceRG.detail.count - 1 do
  begin
    TOBB := TOBPieceRG.detail[Indice];
    if EnHt then
    begin
      XD := XD + TOBB.GetDouble ('PRG_MTHTRGDEV');
      XP := XP + TOBB.GetDouble ('PRG_MTHTRG');
    end else
    begin
      if Reliquat then
      begin
        if (TOBB.GetString ('PRG_NUMCAUTION') <> '') then
        begin
          MtRGResiduelle := TOBB.GetDouble ('PRG_MTTTCRG') +
                            TOBB.GetDouble ('CAUTIONUTIL') -
                            TOBB.GetDouble ('CAUTIONAPRES') -
                            TOBB.GetDouble ('PRG_CAUTIONMT');
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
          XP := XP + MtRGResiduelle;
          //
          MtRGResiduelle := TOBB.GetDouble ('PRG_MTTTCRGDEV') +
                            TOBB.GetDouble ('CAUTIONUTILDEV') -
                            TOBB.GetDouble ('CAUTIONAPRESDEV') -
                            TOBB.GetDouble ('PRG_CAUTIONMTDEV');
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
          XD := XD + MtRGResiduelle;
        end else
        begin
          XP := XP + TOBB.GetDouble ('PRG_MTTTCRG');
          XD := XD + TOBB.GetDouble ('PRG_MTTTCRGDEV');
          if XD > 0 then NumCaution := '';
        end;
      end else
      begin
        XP := XP + TOBB.GetDouble ('PRG_MTTTCRG');
        XD := XD + TOBB.GetDouble ('PRG_MTTTCRGDEV');
        if (TOBB.GetDouble ('PRG_CAUTIONMT') > 0) then
        begin
          MtRGResiduelle := TOBB.GetDouble ('PRG_MTTTCRG') +
                            TOBB.GetDouble ('CAUTIONUTIL') -
                            TOBB.GetDouble ('CAUTIONAPRES') -
                            TOBB.GetDouble ('PRG_CAUTIONMT');
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
          Residu := Residu + MtRGResiduelle;
          //
          MtRGResiduelle := TOBB.GetDouble ('PRG_MTTTCRGDEV') +
                            TOBB.GetDouble ('CAUTIONUTILDEV') -
                            TOBB.GetDouble ('CAUTIONAPRESDEV') -
                            TOBB.GetDouble ('PRG_CAUTIONMTDEV');
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
          ResiduDev := ResiduDev + MtRGResiduelle;
        end else
        begin
          if XD > 0 then NumCaution := '';
        end;
      end;
    end;
  end;
end;

procedure GetRg (TOBpieceRG : TOB; EnHt : boolean; Reliquat : boolean;  var Xp,Xd : double ; var NumCaution : string; Intervenant : string=''; PourTous : boolean=true);
var Indice : integer;
		TOBB  : TOB;
    Sens : Integer;
    MtRGResiduelle,ResiduDev,Residu,mtttcRG,CautionUsed,CautionApres,MtCaution,Ratio : double;
begin
  NumCaution := '1111111';
  Sens := 1;
  Xp := 0; XD := 0;
  residu := 0; ResiDuDev := 0;
  if TOBpieceRG = nil then exit;
  for Indice := 0 to TOBPieceRG.detail.count - 1 do
  begin
    TOBB := TOBPieceRG.detail[Indice];
    if (not PourTous) and (TOBB.getString('PRG_FOURN') <> Intervenant) then continue;
    ratio := 1.0;
    if EnHt then
    begin
      if TOBB.GetDouble ('PRG_MTHTRGDEV') <> 0 then Ratio := TOBB.GetDouble ('PRG_MTTTCRGDEV')/ TOBB.GetDouble ('PRG_MTHTRGDEV');
    end;
      if Reliquat then
      begin
        if (TOBB.GetString ('PRG_NUMCAUTION') <> '') then
        begin
		      // gestion de la RG négative --->  rahhhhhh

          if (TOBB.GetDouble ('PRG_MTTTCRG') < 0)  then
          begin
            Sens := -1;
          end;
          //
          mtttcRG := Abs(TOBB.GetDouble ('PRG_MTTTCRG'));
          CautionUsed := Abs(TOBB.GetDouble ('CAUTIONUTIL'));
          CautionApres := Abs(TOBB.GetDouble ('CAUTIONAPRES'));
          MtCaution := Abs(TOBB.GetDouble ('PRG_CAUTIONMT'));
        MtRGResiduelle := MtttcRG + CautionUsed - MtCaution{- CautionApres};
          //
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
        XP := XP + Arrondi(MtRGResiduelle * sens / ratio,V_PGI.OkDecV);
          //
          //
          mtttcRG := Abs(TOBB.GetDouble ('PRG_MTTTCRGDEV'));
          CautionUsed := Abs(TOBB.GetDouble ('CAUTIONUTILDEV'));
          CautionApres := Abs(TOBB.GetDouble ('CAUTIONAPRESDEV'));
          MtCaution := Abs(TOBB.GetDouble ('PRG_CAUTIONMTDEV'));
        MtRGResiduelle := MtttcRG + CautionUsed  - MtCaution{- CautionApres};
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
        XD := XD + Arrondi(MtRGResiduelle * sens / ratio,V_PGI.okdecV ) ;
        end else
        begin
        if EnHt then
        begin
          XD := XD + TOBB.GetDouble ('PRG_MTHTRGDEV');
          XP := XP + TOBB.GetDouble ('PRG_MTHTRG');
        end else
        begin
          XP := XP + TOBB.GetDouble ('PRG_MTTTCRG');
          XD := XD + TOBB.GetDouble ('PRG_MTTTCRGDEV');
        end;
          if XD <> 0 then NumCaution := '';
        end;
      end else
      begin
      if EnHt then
      begin
        XD := XD + TOBB.GetDouble ('PRG_MTHTRGDEV');
        XP := XP + TOBB.GetDouble ('PRG_MTHTRG');
      end else
      begin
        XP := XP + TOBB.GetDouble ('PRG_MTTTCRG');
        XD := XD + TOBB.GetDouble ('PRG_MTTTCRGDEV');
      end;
        if (TOBB.GetDouble ('PRG_CAUTIONMT') > 0) then
        begin
        	sens := 1;
          if (TOBB.GetDouble ('PRG_CAUTIONMT') < 0) and
						 (TOBB.GetDouble ('PRG_MTTTCRG') < 0)  then
          begin
            Sens := -1;
          end;
          //
          mtttcRG := Abs(TOBB.GetDouble ('PRG_MTTTCRG'));
          CautionUsed := Abs(TOBB.GetDouble ('CAUTIONUTIL'));
          CautionApres := Abs(TOBB.GetDouble ('CAUTIONAPRES'));
          MtCaution := Abs(TOBB.GetDouble ('PRG_CAUTIONMT'));
          MtRGResiduelle := MtttcRG + CautionUsed - CautionApres - MtCaution;

          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
        Residu := Residu + arrondi(MtRGResiduelle * sens / ratio,V_PGI.okdecV);
          //
          Sens := 1;
          if (TOBB.GetDouble ('PRG_CAUTIONMTDEV') < 0) and
						 (TOBB.GetDouble ('PRG_MTTTCRGDEV') < 0)  then
          begin
            Sens := -1;
          end;
          mtttcRG := Abs(TOBB.GetDouble ('PRG_MTTTCRGDEV'));
          CautionUsed := Abs(TOBB.GetDouble ('CAUTIONUTILDEV'));
          CautionApres := Abs(TOBB.GetDouble ('CAUTIONAPRESDEV'));
          MtCaution := Abs(TOBB.GetDouble ('PRG_CAUTIONMTDEV'));
          MtRGResiduelle := MtttcRG + CautionUsed - CautionApres - MtCaution;
          if MtRGResiduelle < 0 then MtRGResiduelle := 0;
          if MtRGResiduelle > 0 then NumCaution := '';
        ResiduDev := ResiduDev + Arrondi(MtRGResiduelle * sens / ratio,V_PGI.okdecV);
        end else
        begin
          if XD <> 0 then NumCaution := '';
        end;
      end;
    end;
end;


procedure GetTVARg (TOBpieceRG : TOB; EnHt : boolean; Reliquat : boolean;  var Xp,Xd : double ; Intervenant : string=''; PourTous : boolean=true;QuePourHT : boolean=false);
var Indice : integer;
		TOBB  : TOB;
    Sens : Integer;
    MtRGResiduelle,mtttcRG,CautionUsed,CautionApres,MtCaution,Ratio,RatioDev,TvaReste : double;
begin
  Sens := 1;
  Xp := 0; XD := 0;
  for Indice := 0 to TOBPieceRG.detail.count - 1 do
  begin
    TOBB := TOBPieceRG.detail[Indice];
    if (QuePourHt) and (TOBB.getString('PRG_TYPERG')<>'HT') then continue; 
    if (not PourTous) and (TOBB.getString('PRG_FOURN') <> Intervenant) then continue;
    ratio := 1.0;
    ratioDev := 1.0;
    if TOBB.GetDouble ('PRG_MTHTRGDEV') <> 0 then Ratio := TOBB.GetDouble ('PRG_MTTTCRGDEV')/ TOBB.GetDouble ('PRG_MTHTRGDEV');
    if TOBB.GetDouble ('PRG_MTTTCRG') <> 0 then RatioDev := TOBB.GetDouble ('PRG_MTTTCRGDEV')/ TOBB.GetDouble ('PRG_MTTTCRG');
    if (TOBB.GetString ('PRG_NUMCAUTION') <> '') and (Reliquat) then
    begin
      // gestion de la RG négative --->  rahhhhhh

      if (TOBB.GetDouble ('PRG_MTTTCRG') < 0)  then
      begin
        Sens := -1;
  end;
      //
      mtttcRG := Abs(TOBB.GetDouble ('PRG_MTTTCRGDEV'));
      CautionUsed := Abs(TOBB.GetDouble ('CAUTIONUTILDEV'));
      CautionApres := Abs(TOBB.GetDouble ('CAUTIONAPRESDEV'));
      MtCaution := Abs(TOBB.GetDouble ('PRG_CAUTIONMTDEV'));
      MtRGResiduelle := MtttcRG + CautionUsed - CautionApres - MtCaution;
      if MtRGResiduelle < 0 then MtRGResiduelle := 0;
      TvaReste := Arrondi(MtRGResiduelle * sens ,V_PGI.okdecV ) - Arrondi(MtRGResiduelle * sens / ratio,V_PGI.okdecV );
      XD := XD +  TvaReste ;
      XP := XP + Arrondi(TvaReste * RatioDev,V_PGI.OkdecV);
    end else
    begin
      XP := XP +  Arrondi(TOBB.GetDouble ('PRG_MTTTCRG')-TOBB.GetDouble ('PRG_MTHTRG'),V_PGI.OkDecV);
      XD := XD + Arrondi(TOBB.GetDouble ('PRG_MTTTCRGDEV')-TOBB.GetDouble ('PRG_MTHTRGDEV'),V_PGI.OkDecV);
    end;
  end;
end;

procedure GetCautionUsedAfter (TOBPiece : TOB; PiecePrec : string;  NumSituation : integer;var CautionApres,CautionApresdev : double ;Intervenant : string = '');
var Sql : string;
		QQ : TQuery;
    TOBC : TOB;
    Indice : integer;
begin
  TOBC := TOB.Create ('LES CAUTION',nil,-1);
	SQL := 'SELECT BST_NATUREPIECE,BST_SOUCHE,BST_NUMEROFAC,BST_NUMEROSIT,'+
  			 'PRG_NUMLIGNE,PRG_FOURN,PRG_MTTTCRGDEV,PRG_MTTTCRG,GL_PIECEPRECEDENTE '+
         'FROM BSITUATIONS '+
         'LEFT JOIN PIECERG ON PRG_NATUREPIECEG=BST_NATUREPIECE AND '+
         'PRG_SOUCHE=BST_SOUCHE AND PRG_NUMERO=BST_NUMEROFAC AND '+
         'PRG_INDICEG=0 LEFT JOIN LIGNE ON GL_NATUREPIECEG=BST_NATUREPIECE AND '+
         'GL_SOUCHE=BST_SOUCHE AND GL_NUMERO=BST_NUMEROFAC AND GL_INDICEG=0 AND Gl_NUMORDRE=PRG_NUMLIGNE '+
         'WHERE BST_SSAFFAIRE="'+TOBPiece.GetValue('GP_AFFAIREDEVIS')+'" AND '+
         'BST_NUMEROSIT > '+IntToStr(NumSituation)+' AND '+
         'BST_VIVANTE ="X" AND '+
         'PRG_FOURN ="'+Intervenant+'" AND '+
         'GL_PIECEPRECEDENTE="'+PiecePrec+'"';
  QQ := OpenSQL(Sql,true,-1,'',true);
  if not QQ.eof then
  begin
		TOBC.LoadDetailDB('BSITUATIONS','','',QQ,false);
    for Indice := 0 to TOBC.detail.count -1 do
    begin
      CautionApres := CautionApres + TOBC.getDouble ('PRG_MTTTCRG');
      CautionApresDEV := CautionApresDEV + TOBC.getDouble ('PRG_MTTTCRGDEV');
    end;
  end;
  TOBC.Free;
end;

procedure GetCautionAlreadyUsed (TOBpiece,TOBL: TOB; PieceOrigine : string;var CautionUSed,CautionUsedDev : double; Intervenant : string =''; NewPiece : boolean=false);
var CD : R_CLEDOC;
		QQ : TQuery;
    TOBSIt : TOB;
    req,Typefacturation : String;
    II : Integer;
begin
	CautionUsed := 0;
  CautionUsedDev := 0;
  if PieceOrigine = '' then exit;
  TOBSIt := TOB.Create ('LES SITUATIONS',nil,-1);
  DecodeRefPiece (PieceOrigine,CD);
  TypeFacturation := RenvoieTypeFact(TOBPiece.GetValue('GP_AFFAIREDEVIS'));
  if Typefacturation='AVA' then
  begin
    if NewPiece then
    begin
      Req := 'SELECT BST_NATUREPIECE AS NATUREPIECE ,BST_SOUCHE AS SOUCHE ,BST_NUMEROFAC AS NUMERO, 0 AS INDICE '+
             'FROM BSITUATIONS WHERE '+
             'BST_SSAFFAIRE="'+TOBpiece.getValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_VIVANTE="X" ORDER BY BST_NUMEROSIT';
    end else
    begin
      Req := 'SELECT BST_NATUREPIECE AS NATUREPIECE ,BST_SOUCHE AS SOUCHE ,BST_NUMEROFAC AS NUMERO, 0 AS INDICE '+
             'FROM BSITUATIONS WHERE '+
             'BST_SSAFFAIRE="'+TOBpiece.getValue('GP_AFFAIREDEVIS')+'" AND '+
             'BST_VIVANTE="X" AND '+
             'BST_NUMEROSIT < '+
             '('+
                'SELECT BST_NUMEROSIT FROM BSITUATIONS WHERE '+
                'BST_NATUREPIECE="'+TOBpiece.getValue('GP_NATUREPIECEG')+'" AND '+
                'BST_SOUCHE="'+TOBpiece.getValue('GP_SOUCHE')+'" AND '+
                'BST_NUMEROFAC='+InttoStr(TOBpiece.getValue('GP_NUMERO'))+
             ') ORDER BY BST_NUMEROSIT';
    end;
    QQ := OpenSQL(req,true,-1,'',true);
    if Not QQ.eof then
    begin
      TOBSIt.LoadDetailDB('LES DOCUMENTS','','',QQ,false);
    end;
    ferme (QQ);
  end else
  begin
    if NewPiece then
    begin
      Req := 'SELECT GP_NATUREPIECEG AS NATUREPIECE ,GP_SOUCHE AS SOUCHE ,GP_NUMERO AS NUMERO, GP_INDICEG AS INDICE '+
             'FROM PIECE WHERE '+
             'GP_NATUREPIECEG IN ("FBP","FBT") AND '+
             'GP_AFFAIREDEVIS="'+TOBpiece.getValue('GP_AFFAIREDEVIS')+'" AND '+
             'GP_VIVANTE="X"';
    end else
    begin
      Req := 'SELECT GP_NATUREPIECEG AS NATUREPIECE ,GP_SOUCHE AS SOUCHE ,GP_NUMERO AS NUMERO, GP_INDICEG AS INDICE '+
             'FROM PIECE WHERE '+
             'GP_NATUREPIECEG IN ("FBP","FBT") AND '+
             'GP_AFFAIREDEVIS="'+TOBpiece.getValue('GP_AFFAIREDEVIS')+'" AND '+
             'GP_DATEPIECE < "'+USDATETIME(TOBPiece.getvalue('GP_DATEPIECE'))+'" AND '+
             'GP_VIVANTE="X"';
    end;
    QQ := OpenSQL(req,true,-1,'',true);
    if Not QQ.eof then
    begin
      TOBSIt.LoadDetailDB('LES DOCUMENTS','','',QQ,false);
    end;
    ferme (QQ);

  end;
  if TOBSIt.Detail.count > 0 then
  begin
    for II := 0 to TOBSIT.detail.count -1 do
    begin
      req := 'SELECT PRG_MTTTCRG,PRG_MTTTCRGDEV FROM PIECERG '+
             'WHERE '+
             'PRG_NATUREPIECEG="'+TOBSit.detail[II].GetString('NATUREPIECE')+'" AND '+
             'PRG_SOUCHE="'+TOBSit.detail[II].GetString('SOUCHE')+'" AND '+
             'PRG_NUMERO='+TOBSit.detail[II].GetString('NUMERO')+' AND '+
             'PRG_INDICEG='+TOBSit.detail[II].GetString('INDICE')+' AND '+
             'PRG_NUMLIGNE='+
             '('+
             'SELECT GL_NUMORDRE FROM LIGNE WHERE '+
             'GL_NATUREPIECEG="'+TOBSit.detail[II].GetString('NATUREPIECE')+'" AND '+
             'GL_SOUCHE="'+TOBSit.detail[II].GetString('SOUCHE')+'" AND '+
             'GL_NUMERO='+TOBSit.detail[II].GetString('NUMERO')+' AND '+
             'GL_INDICEG='+TOBSit.detail[II].GetString('INDICE')+' AND '+
             'GL_PIECEORIGINE="'+PieceOrigine+'"'+
             ')';
      QQ := OpenSQL(req,True,1,'',true);
      if not QQ.eof then
      begin
        CautionUsed := CautionUsed+ QQ.fields[0].asFloat  ;
        CautionUsedDEV := CautionUsedDEV + QQ.fields[1].asFloat;
      end;
      ferme (QQ);
    end;
  end;


  TOBSIt.free;
  (*
  req := 'SELECT PRG_CAUTIONMTU,PRG_CAUTIONMTUDEV FROM PIECERG WHERE '+WherePiece (CD,ttdretenuG,true)+
  							 ' AND PRG_NUMLIGNE='+IntToStr(CD.NumLigne)+ 'AND PRG_FOURN="'+Intervenant+'"';
  TRY
    QQ := OpenSql (Req,true,-1, '', True);
    if not QQ.eof then
    begin
      CautionUsed := QQ.findField('PRG_CAUTIONMTU').asFloat  ;
      CautionUsedDEV := QQ.findField('PRG_CAUTIONMTUDEV').asFloat;
    end;
  FINALLY
  	if assigned(QQ) then ferme (QQ);
  END;
  *)
end;

procedure GetMontantRGReliquat(TOBPieceRg : TOB; var RGRP,RGRD : double; ForceTTC : boolean=false; UniquementsiTTC : boolean=false) ;
var Xp,XD : double;
		indice : integer;
    TOBLoc :TOb;
    Sens, MtTTCRG,MtCaution,CautionUsed,ratio : double;
    //MtTTCRG : double;
    MtHTRG  : Double;
begin
  RGRP := 0 ;RGRD := 0;
  if TOBPieceRg = nil then Exit;
  for Indice := 0 to TOBPieceRG.detail.count -1 do
  begin
  	TOBLoc := TOBPieceRG.detail[Indice];
    //FV1 : 02/10/2015 - FS#1726 - TEAM RESEAUX - pièce inutilisable avec message 'opération en virgule flottante incorrecte'
    MtTTCRG := TOBLOC.GetValue('PRG_MTTTCRG');
    MtHTRG  := TOBLOC.GetValue('PRG_MTHTRG');
    if (MtTTCRG = 0) Or (MtHTRG=0) then Continue;

    Ratio := 1;
  	if TOBLOC = nil then exit;
    if (TOBLoc.getValue('PRG_TYPERG')='HT') and (UniquementsiTTC) then continue;
    if (TOBLoc.getValue('PRG_TYPERG')='HT') and (not forceTTC) then
    begin
      XP:= TOBLOC.GetValue('PRG_MTHTRG');
      XD:= TOBLOC.GetValue('PRG_MTHTRGDEV');
    end else
    begin
      XP:= TOBLOC.GetValue('PRG_MTTTCRG');
      XD:= TOBLOC.GetValue('PRG_MTTTCRGDEV');
    end;

    if (TOBLoc.getValue('PRG_TYPERG')='HT') and (not forceTTC) then Ratio := TOBLOC.GetValue('PRG_MTTTCRG') /TOBLOC.GetValue('PRG_MTHTRG');
    if (TOBLoc.GetValue('PRG_NUMCAUTION')<> '') and (Pos(TOBLoc.GetValue('PRG_NATUREPIECEG'),'ABT;ABP')=0) then
    begin
      // gestion de la RG négative --->  rahhhhhh
      if TOBLOC.getValue('PRG_MTTTCRG') < 0 then Sens := -1 else Sens := 1;
      MtTTCRG := Abs(TOBLOC.GetValue('PRG_MTTTCRG'));
      MtCaution := Abs(TOBLOC.getValue('PRG_CAUTIONMT'));
      CautionUsed := Abs(TOBLOC.getValue('CAUTIONUTIL'));
      XP := (MtTTCRG - (MtCaution - CautionUsed));
      if XP < 0 then Xp := 0;
      Xp := Arrondi(XP * Sens / ratio,V_PGI.okdecV);
      //
      MtTTCRG := Abs(TOBLOC.GetValue('PRG_MTTTCRGDEV'));
      MtCaution := Abs(TOBLOC.getValue('PRG_CAUTIONMTDEV'));
      CautionUsed := Abs(TOBLOC.getValue('CAUTIONUTILDEV'));
      XD := (MtTTCRG - (MtCaution - CautionUsed));
      if XD < 0 then XD := 0;
      XD := Arrondi(XD * Sens / ratio,V_PGI.okdecV);
    end;
    RGRP := RGRP + XP;
    RGRD := RGRD + XD;
  end;
end;

procedure GetReliquatCaution (TOBPieceRg : TOB; var RQ,RQDev : double) ;
var indice : integer;
		TOBPRG : TOB;
begin
	RQ := 0;
  RQDev := 0;
	for Indice := 0 to TOBPIeceRG.detail.count -1 do
  begin
  	TOBPRg := TOBPieceRG.detail[Indice];
    if Indice = 0 then
    begin
			if (TOBPrg.GetValue('PRG_NUMCAUTION')<> '') then
      begin
    		RQ := RQ + TOBPrg.getValue('PRG_CAUTIONMT') - TOBPrg.getValue('CAUTIONUTIL');
    		RQDev := RQDev + TOBPrg.getValue('PRG_CAUTIONMTDEV') - TOBPrg.getValue('CAUTIONUTILDEV');
      end;
    end;
  end;
end;

procedure GetValueRgCot (TOBPieceRG: TOB; var RGRP,RGRD :double ;Reliquat:boolean);
var MyCautionrest : double;
		Indice : integer;
begin
  RGRP := 0 ;RGRD := 0;
  if TOBPieceRG = nil then exit;
  for Indice := 0 to TOBPieceRG.detail.count - 1 do
  begin
    RGRP:=RGRP + TOBPieceRG.detail[Indice].GetValue('PRG_MTTTCRG');
    RGRD:=RGRD + TOBPieceRG.detail[Indice].GetValue('PRG_MTTTCRGDEV');
  end;
  if (Pos(TOBPieceRG.getValue('PRG_NATUREPIECEG'),'ABT;ABP')=0) Then
  begin
    if RGRP < 0 then RGRP := 0;
    if RGRD < 0 then RGRD := 0;
  end;
end;


procedure GetValueRg (TOBLOC: TOB; var RGRP,RGRD :double ;Reliquat:boolean);
var MyCautionrest,CautionUsed,CautionUsedDev : double;
begin
  RGRP := 0 ;RGRD := 0;
  CautionUsed := 0; CautionUsedDev := 0 ;
  if TOBLOC = nil then exit;
  RGRP:=TOBLOC.GetValue('PRG_MTTTCRG');
  RGRD:=TOBLOC.GetValue('PRG_MTTTCRGDEV');
  if (TOBLoc.getValue('PRG_NATUREPIECEG') <> 'ABT') and (TOBLoc.getValue('PRG_NATUREPIECEG') <> 'ABP') Then
  begin
    if RGRP < 0 then RGRP := 0;
    if RGRD < 0 then RGRD := 0;
  end;
end;

procedure GetCumulTaxeByIndiceCot (TOBPIECE,TOBPIECERG,TOBBasesRG:TOB; var TP,TD :double);
var TOBLOC : TOB;
    XD,XP : double;
    RD,RP : double;
    Indice:integer;
    Ratio : double;
    DEV : Rdevise;
begin
  TP:=0;TD:=0;
  if TOBPIeceRG=nil then exit;
  DEV.Code:=TOBPiece.getValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
  Indice:=TOBPieceRG.detail[0].getValue('INDICERG');
  // recupere le reliquat de RG par ligne
  GetValueRgCot  (TOBPieceRG,RP,RD,true);
  GetValueRgCot (TOBPieceRG,XP,XD,false);
  if XD <> 0 then Ratio := RD / XD else Ratio := 0;
  // --
  if TOBBasesRG=nil then exit;
  TOBLOC:=TOBBasesRG.findfirst(['INDICERG'],[Indice],true);
  While TOBLOC <> nil do
  begin
    GetLaTaxeRG (TOBLOC,XD,XP);
    TP := TP+(XP*Ratio);
    TD := TD+(XD*Ratio);
    TOBLOC:=TOBBasesRG.findnext(['INDICERG'],[Indice],true);
  end;
  TD:=Arrondi (TD,DEV.Decimale );
  TP:=Arrondi (TP,V_PGI.okDecv );
end;


procedure GetCumulTaxeByIndice (TOBPIECE,TOBPIECERG,TOBBasesRG:TOB; var TP,TD :double);
var TOBLOC : TOB;
    XD,XP : double;
    RD,RP : double;
    Indice:integer;
    Ratio : double;
    DEV : Rdevise;
begin
TP:=0;TD:=0;
if TOBPIeceRG=nil then exit;
DEV.Code:=TOBPiece.getValue('GP_DEVISE') ; GetInfosDevise(DEV) ;
Indice:=TOBPieceRG.getValue('INDICERG');
// recupere le reliquat de RG par ligne
GetValueRg (TOBPieceRG,RP,RD,true);
GetValueRg (TOBPieceRG,XP,XD,false);
if XD <> 0 then Ratio := RD / XD else Ratio := 0;
// --
if TOBBasesRG=nil then exit;
TOBLOC:=TOBBasesRG.findfirst(['INDICERG'],[Indice],true);
While TOBLOC <> nil do
   begin
   GetLaTaxeRG (TOBLOC,XD,XP);
   TP := TP+(XP*Ratio);
   TD := TD+(XD*Ratio);
   TOBLOC:=TOBBasesRG.findnext(['INDICERG'],[Indice],true);
   end;
TD:=Arrondi (TD,DEV.Decimale );
TP:=Arrondi (TP,V_PGI.okDecv );
end;

procedure ReAjusteMontants ( TOBLOC: TOB;DEV:Rdevise);
begin
if TOBLOC.NomTable = 'PIEDBASERG' then
   begin
   CoordinateMont (TOBLOC,'PBR_VALEURDEV','PBR_VALEURTAXE',DEV);
   CoordinateMont (TOBLOC,'PBR_BASEDEV','PBR_BASETAXE',DEV);
   end else
   if TOBLOC.NomTable = 'PIECERG' then
     begin
     CoordinateMont (TOBLOC,'PRG_MTHTRGDEV','PRG_MTHTRG',DEV);
     CoordinateMont (TOBLOC,'PRG_MTTTCRGDEV','PRG_MTTTCRG',DEV);
     end;
end;

procedure VideBasesRGLigne (TOBBasesRG: TOB;Ligne : integer);
var TOBL : TOB;
begin
if TOBBasesRG = nil then exit;
TOBL := TOBBasesRG.findfirst (['INDICERG'],[Ligne],true);
while TOBL <> nil do
   begin
   TOBL.free;
   TOBL := TOBBasesRG.findnext (['INDICERG'],[Ligne],true);
   end;
end;

procedure CalculeLigneBasesRG(TOBPIECERG,TOBB,TOBBRG: TOB;Taux:double;Dev:Rdevise;IndiceRg:Integer);
begin
TOBBRG.PutValue ('PBR_NATUREPIECEG',TOBB.GetValue('GPB_NATUREPIECEG'));
TOBBRG.PutValue ('PBR_DATEPIECE',TOBB.GetValue('GPB_DATEPIECE'));
TOBBRG.PutValue ('PBR_SOUCHE',TOBB.GetValue('GPB_SOUCHE'));
TOBBRG.PutValue ('PBR_NUMERO',TOBB.GetValue('GPB_NUMERO'));
TOBBRG.PutValue ('PBR_INDICEG',TOBB.GetValue('GPB_INDICEG'));
TOBBRG.PutValue ('PBR_FOURN',TOBB.GetValue('GPB_FOURN'));
TOBBRG.Putvalue('INDICERG',IndiceRG);
TOBBRG.PutValue ('PBR_CATEGORIETAXE',TOBB.GetValue('GPB_CATEGORIETAXE'));
TOBBRG.PutValue ('PBR_FAMILLETAXE',TOBB.GetValue('GPB_FAMILLETAXE'));
TOBBRG.PutValue ('PBR_TAUXTAXE',TOBB.GetValue('GPB_TAUXTAXE'));
TOBBRG.PutValue ('PBR_TAUXDEV',DEV.Taux );
TOBBRG.PutValue ('PBR_COTATION',dev.Cotation);
TOBBRG.PutValue ('PBR_DATETAUXDEV',DEV.DateTaux);
TOBBRG.PutValue ('PBR_DEVISE',DEV.Code );
TOBBRG.PutValue ('PBR_SAISIECONTRE',TOBB.GetValue('GPB_SAISIECONTRE'));
TOBBRG.PutValue ('PBR_BASEDEV',arrondi(TOBB.GetValue('GPB_BASEDEV')*Taux,DEV.Decimale  ));
TOBBRG.PutValue ('PBR_VALEURDEV',arrondi(TOBBRG.GetValue('PBR_BASEDEV')*(TOBBRG.GetValue('PBR_TAUXTAXE')/100),DEV.Decimale ));
ReAjusteMontants (TOBBRG,DEV);
end;

procedure DeduitsPortRg (TOBBasesRg,TOBporcs : TOB; Taux : double; DEv : RDevise; IndiceRG : integer);
var Indice : integer;
		TOBP, TOBBRG : TOB;
    MtPort : double;
begin
	if TOBPorcs = nil then exit;
	for Indice := 0 to TOBPorcs.detail.count -1 do
  begin
  	TOBP := TOBporcs.detail[Indice];
    if TOBP.GetValue('GPT_FRAISREPARTIS') = 'X' then Continue; {ce sont des frais donc pas inclus dans le calcul}
    if TOBP.getboolean('GPT_RETENUEDIVERSE') then Continue; {ce sont des retenues diverses donc pas inclus dans le calcul}
    MtPort := Arrondi(TOBP.GetValue('GPT_TOTALHTDEV')*taux,DEV.decimale);
    TOBBRG := TOBBasesRG.findFirst(['INDICERG','PBR_CATEGORIETAXE','PBR_FAMILLETAXE'],[indiceRG,'TX1',TOBP.getValue('GPT_FAMILLETAXE1')],true);
    repeat
    	if TOBBRG <> nil then
      begin
      	TOBBRG.PutValue ('PBR_BASEDEV',arrondi(TOBBRG.GetValue('PBR_BASEDEV')-MtPort,DEV.decimale ));
        TOBBRG.PutValue ('PBR_VALEURDEV',arrondi(TOBBRG.GetValue('PBR_BASEDEV')*(TOBBRG.GetValue('PBR_TAUXTAXE')/100),DEV.Decimale ));
        ReAjusteMontants (TOBBRG,DEV);
      end;
    	TOBBRG := TOBBasesRG.FindNext (['INDICERG','PBR_CATEGORIETAXE','PBR_FAMILLETAXE'],[IndiceRg,'TX1',TOBP.getValue('GPT_FAMILLETAXE1')],true);
    until TOBBRG = nil;
  end;
end;

// Calcul des elements TVA
procedure CalculeTvaRg (TOBPieceRg,TOBBases,TOBBasesRg,TOBporcs : TOB;DEv:Rdevise;Taux:double;IndiceRG:Integer;Var TotTaxes:double);
var MTTva,MtTvaDev : double;
    Indice : integer;
    TOBB,TOBBRG : TOB;
begin
  if TOBPieceRG = nil then exit;
  if TOBBases = nil then exit;
  TotTaxes:=0;
  VideBasesRGLigne (TOBBasesRg,IndiceRG);
  MTTva:=0; MTTvaDev :=0;
  for Indice:= 0 to TOBBASES.detail.count -1 do
    begin
    TOBB:=TOBBases.detail[Indice];
    TOBBRG := TOB.Create('PIEDBASERG',TOBBASESRG,-1);
    TOBBRG.addChampsupvaleur ('INDICERG',IndiceRG);
    CalculeLigneBasesRG(TOBPIECERG,TOBB,TOBBRG,Taux,DEV,IndiceRg);
  end;
  // Deduction des ports et frais
//  DeduitsPortRg (TOBBasesRg,TOBporcs,Taux,DEV,IndiceRG);
  // recalcul des Tvas sur RG
  for Indice:= 0 to TOBBASESRG.detail.count -1 do
  begin
    TOBBRG := TOBBASESRG.detail[indice];
    if (TOBBRG.getValue ('INDICERG')<>IndiceRg) then continue;
    MTTva := MTTva+TOBBRG.GetValue('PBR_VALEURTAXE');
    MTTvaDev := MTTvaDev+TOBBRg.GetValue('PBR_VALEURDEV');
  end;
	TotTaxes := MtTvaDev;
end;

function CalculPort (EnHt: boolean;TOBPorcs : TOB) : double;
var i : integer;
begin
  result := 0;
  if TOBPorcs <> nil then
  begin
    for i := 0 to TOBPorcs.Detail.Count - 1 do
    begin
      if (TOBPorcs.Detail[i].GetValue('GPT_FRAISREPARTIS') = 'X') then continue;
      if (TOBPorcs.Detail[i].Getboolean('GPT_RETENUEDIVERSE')) then continue;
    if EnHt then
    begin
        Result := Result + TOBPorcs.Detail[i].GetValue('GPT_TOTALHTDEV');
    end else
    begin
        Result := Result + TOBPorcs.Detail[i].GetValue('GPT_TOTALTTCDEV');
    end;
  end;
  end;
  Result := Arrondi(Result, 6);
end;

function GetBaseCoTrait(TOBBases : TOB; Intervenant : string;EnHt : boolean) : double;
var TOBT : TOB;
		Indice : integer;
begin

	result := 0;

  if TOBBases = nil then exit;

  if TobBases <> nil then
  begin
    for Indice := 0 to TOBBases.detail.count -1 do
    begin
      if EnHt then
      begin
        result := result + TOBBases.detail[Indice].getDouble('GPB_BASEDEV')
      end else
      begin
        result := result + TOBBases.detail[Indice].getDouble('GPB_BASEDEV')+TOBBases.detail[Indice].getDouble('GPB_VALEURDEV');
      end;
    end;
  end;


end;

function GetCumulPorcs(TOBPorcsCot: TOB; IndicePort : integer; Ht : boolean) : double;
var Indice : integer;
begin
  result := 0;
  if TOBPorcsCot <> nil then
  begin
    for Indice := 0 to TOBPorcsCot.detail.count -1 do
    begin
      if Ht then
      begin
        result := result + TOBPorcsCot.detail[Indice].detail[IndicePort].GetDouble ('GPT_TOTALHTDEV');
      end else
      begin
        result := result + TOBPorcsCot.detail[Indice].detail[IndicePort].GetDouble('GPT_TOTALTTCDEV');
      end;
    end;
  end;
end;

procedure  AffecteEcart (TOBPorcsCot : TOB;IndiceInterv,Indice : integer; Ecart: double ;Ht : boolean);
var TOBP : TOB;
begin

  if TOBPorcsCot <> nil then
  begin
    TOBP := TOBPorcsCot.detail[IndiceInterv];
    if Ht then
    begin
      TOBP.Detail[Indice].SetDouble ('GPT_TOTALHTDEV',TOBP.Detail[Indice].GetDouble ('GPT_TOTALHTDEV')+Ecart);
    end else
    begin
      TOBP.Detail[Indice].SetDouble ('GPT_TOTALTTCDEV',TOBP.Detail[Indice].GetDouble ('GPT_TOTALTTCDEV')+Ecart);
    end;
  end;

end;

procedure ProratisePorcsCot (TOBPorcs,TOBPorcsCot,TOBPiece,TOBpieceTrait : TOB; DEV : rdevise);
var Indice,I,IndiceMaxVal : integer;
		XPorcHt,XPorcTTC,XX : double;
    Prorata,MaxVal,Ecart,Base : double;
    TOBPCT,TOBPENT,TOBDENT : TOB;
    Intervenant:  string;
begin
  XPorcHt := CalculPort (True,TOBPorcs);
  XPorcTTC := CalculPort (False,TOBPorcs);
  if XPorcHt = 0 then exit;
  maxVal := 0; IndiceMaxVal := 0;
  //
	for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
    Intervenant := TOBPieceTrait.detail[Indice].GetString('BPE_FOURNISSEUR');
    Base := GetBaseCoTrait(TOBPieceTrait,Intervenant,true); // Base HT pour l'intervenant

    TOBPCt := TOB.Create ('PORT COTRAIT',TOBPorcsCot,-1);
    TOBPCT.AddChampSupValeur('FOURN',Intervenant);
    TOBPCT.AddChampSupValeur('TYPEPAIE',TOBPieceTrait.detail[Indice].GetString('TYPEPAIE'));
    TOBPCT.Dupliquer(TOBPorcs,true,true);
    //
    if Base > Maxval Then     // pour les reajustements
    begin
      MaxVal := Base;
      IndiceMaxVal := Indice;
    end;
    Prorata := Base / TOBPiece.getdouble('GP_TOTALHTDEV');

    for I := 0 to TOBPCT.detail.count -1 do
    begin
      if TOBPCT.Detail[i].GetValue('GPT_FRAISREPARTIS') = 'X' then continue;
      if (TOBPCT.Detail[i].Getboolean('GPT_RETENUEDIVERSE')) then continue;
      //
      XX := TOBPCT.detail[I].GetDouble ('GPT_TOTALHTDEV');
      TOBPCT.detail[I].SetDouble ('GPT_TOTALHTDEV',Arrondi(XX*Prorata,V_PGI.okdecV));
      //
      XX := TOBPCT.detail[Indice].GetDouble ('GPT_TOTALTTCDEV');
      TOBPCT.detail[I].SetDouble ('GPT_TOTALTTCDEV',Arrondi(XX*Prorata,V_PGI.okdecV));
    end;
  end;
  // On cumule les ports sous traitants payés par l'entreprise sur l'entreprise
  I := 0;
  (*
  TOBPENT := TOBPorcsCot.FindFirst(['FOURN'],[''],true);
  repeat
    TOBPCT := TOBPorcsCot.detail[I];
    if TOBPCT.GetString('TYPEPAIE')='002' then
    begin
      if TOBPENT = nil then
      begin
        TOBPCT.SetString('FOURN','');
        TOBPENT := TOBPCT;
        inc(I);
      end else
      begin
        for Indice := 0 to TOBPCT.detail.count -1 do
        begin
          XPorcHt := TOBPCT.detail[Indice].GetDouble('GPT_TOTALHTDEV');
          XPorcTTC  := TOBPCT.detail[Indice].GetDouble('GPT_TOTALTTCDEV');
          TOBDENT := TOBPENT.findFirst(['GPT_NUMPORT'],[TOBPCT.detail[Indice].GetInteger('GPT_NUMPORT')],true);
          if TOBDent <> nil then
          begin
            TOBDENT.setDouble('GPT_TOTALHTDEV',TOBDENT.GetDouble('GPT_TOTALHTDEV')+XPorcHt);
            TOBDENT.setDouble('GPT_TOTALTTCDEV',TOBDENT.GetDouble('GPT_TOTALTTCDEV')+XPorcTTC);
          end;
        end;
        TOBPCT.free;
      end;
    end;
  until I>= TOBPorcsCot.detail.count -1;
  *)
  //
  for Indice := 0 to TOBPorcs.detail.count -1 do
  begin
  	if TOBPorcs.Detail[Indice].GetBoolean('GPT_FRAISREPARTIS') then continue;
  	if TOBPorcs.Detail[Indice].GetBoolean('GPT_RETENUEDIVERSE') then continue;
    XX := GetCumulPorcs(TOBPorcsCot,Indice,true);
    ecart := Arrondi(TOBPorcs.Detail[Indice].GetDouble('GPT_TOTALHTDEV') - XX,V_PGI.okdecV);
    if Ecart <> 0 then
    begin
      AffecteEcart (TOBPorcsCot,IndiceMaxVal,Indice,Ecart,True);
    end;
    //
    XX := GetCumulPorcs(TOBPorcsCot,Indice,false);
    ecart := Arrondi(TOBPorcs.Detail[Indice].GetDouble('GPT_TOTALTTCDEV') - XX,V_PGI.okdecV);
    if Ecart <> 0 then
    begin
      AffecteEcart (TOBPorcsCot,IndiceMaxVal,Indice,Ecart,false);
    end;
  end;
end;

function FindBasesIntervenant (TOBBasesCot : TOB; Intervenant : string) : TOB;
var Indice : integer;
begin
  result := nil;
	For Indice := 0 to TOBBasesCot.detail.count -1 do
  begin
    if TOBBasesCot.detail[Indice].GetString('FOURN')=Intervenant then
    begin
      Result := TOBBasesCot.detail[Indice];
      break;
    end;
  end;
end;

function FindPortIntervenant (TOBPOrcsCot: TOB; Intervenant : string) : TOB;
var indice : integer;
begin
	result := nil;
  for Indice := 0 to TOBPOrcsCot.detail.count -1 do
  begin
    if TOBPOrcsCot.detail[Indice].GetValue('FOURN')=Intervenant then
    begin
      result := TOBPOrcsCot.detail[Indice];
      break;
    end;
  end;
end;

procedure ConstituelesBasesCot (TOBPieceTrait,TOBBasesCot,TOBBASES : TOB);
var Indice,I : integer;
		TOBB,TOBBD,TOBPT,TOBBE,TOBBED : TOB;
    Intervenant : string;
begin
  TOBBasesCot.clearDetail;
	for Indice := 0 to TOBBAses.detail.count -1 do
  begin
		Intervenant := TOBBAses.detail[Indice].GetString('GPB_FOURN');
    TOBB := FindBasesIntervenant (TOBBasesCot,Intervenant);
    if TOBB = nil then
    begin
      TOBB := TOB.Create ('LES BASES INTERV',TOBBasesCot,-1);
      TOBB.AddChampSupValeur ('FOURN',Intervenant);
    end;
    TOBBD := TOB.Create ('PIEDBASE',TOBB,-1);
    TOBBD.Dupliquer(TOBBASES.detail[Indice],false,true);
  end;
  TOBBE := TOBBasesCot.findFirst(['FOURN'],[''],true);
  (*
  Indice := 0;
  repeat
    TOBB := TOBBasesCot.detail[Indice];
    TOBPT := TOBpieceTrait.findFirst(['BPE_FOURNISSEUR'],[TOBB.GetString('FOURN')],true);
    if TOBPT <> nil then
    begin
      if TOBPT.getString('TYPEPAIE')='002' then
      begin
        // Sous traitant payé par l'entreprise
        if TOBBE = nil then
        begin
          TOBB.SetString('FOURN','');
          for I := 0 to TOBB.detail.count -1 do
          begin
            TOBB.detail[Indice].SetString('GPB_FOURN','');
          end;
      		inc(Indice);
        end else
        begin
          I := 0;
          repeat
            TOBBD := TOBB.detail[I];
            TOBBED := TOBBE.findFirst(['GPB_CATEGORIETAXE','GPB_FAMILLETAXE'],
            													[TOBBD.GetString('GPB_CATEGORIETAXE'),TOBBD.GetString('GPB_FAMILLETAXE')],true);
            if TOBBED = nil then
            begin
            	TOBBD.ChangeParent (TOBBE,-1);
						end else
            begin
							TOBBED.SetDouble('GPB_BASETAXE',TOBBED.GetDouble('GPB_BASETAXE')+TOBBD.GetDouble('GPB_BASETAXE'));
							TOBBED.SetDouble('GPB_BASEDEV',TOBBED.GetDouble('GPB_BASEDEV')+TOBBD.GetDouble('GPB_BASEDEV'));
							TOBBED.SetDouble('GPB_VALEURTAXE',TOBBED.GetDouble('GPB_VALEURTAXE')+TOBBD.GetDouble('GPB_VALEURTAXE'));
							TOBBED.SetDouble('GPB_VALEURDEV',TOBBED.GetDouble('GPB_VALEURDEV')+TOBBD.GetDouble('GPB_VALEURDEV'));
              inc(I);
            end;
          until I>=TOBB.detail.count -1;
          TOBB.free;
        end;
      end else
      begin
        inc(Indice);
      end;
    end else
    begin
      Inc(Indice);
    end;
  until Indice >= TOBBasesCot.detail.count ;
  *)
end;

procedure CalculeRGSimpleCotrait (TOBPorcs,TOBPIECE,TOBPieceRG,TOBBASES,TOBBASESRG,TOBPieceTrait: TOB;DEV : RDevise);
var Indice,IndiceRg,IInd : integer;
		TOBPOrcsCot,TOBBasesCot,TOBRG,TOBPortLOC,TOBasesLOC,TOBBRLOC,TOBBasesRGLoc : TOB;
    TauxRg,Base,XPorcHt,XPorcTTC,ValInterm,MontantTaxes : double;
    TypeRG : string;
begin
  TOBBasesRG.ClearDetail;
  TOBPorcsCot := TOB.Create ('LES PORTS COT',nil,-1);
  TOBBasesCot := TOB.Create ('LES BASES COT',nil,-1);
  TOBBasesRGLoc := TOB.Create ('LES BASES RG COT',nil,-1);
  TRY
    //
    ProratisePorcsCot (TOBPorcs,TOBPorcsCot,TOBPiece,TOBpieceTrait,DEV);
    ConstituelesBasesCot (TOBPieceTrait,TOBBasesCot,TOBBASES);
    //
		ConstituePieceRg (TOBPIECE,TOBPieceRG,TOBPieceTrait);
    for Indice := 0 to TOBPieceRG.detail.count -1 do
    begin
      TOBRG := TOBPieceRG.detail[Indice];
      TOBRG.putvalue('PRG_MTHTRGDEV',0);
      TOBRG.putvalue('PRG_MTTTCRGDEV',0);
      TOBRG.putvalue('PRG_MTTTCRGDEV',0);
      TOBRG.putvalue('PRG_MTHTRGDEV',0);
      TOBRG.putvalue('PRG_MTHTRG',0);
      TOBRG.putvalue('PRG_MTTTCRG',0);
      //
      if (TOBRG.GetValue('PRG_TAUXRG') <> 0) and (TOBRG.GetBoolean('PRG_APPLICABLE')) then
      begin
        // Taux retenue
        TAuxRg := TOBRG.GetValue('PRG_TAUXRG');
        TypeRg := TOBRG.GetValue('PRG_TYPERG');
        // recup TOB des port et bases de taxes sur intervenant
        TOBPortLOC := FindPortIntervenant (TOBPOrcsCot,TOBRG.GetString('PRG_FOURN'));
        TOBasesLOC := FindBasesIntervenant(TOBBasesCot,TOBRG.GetString('PRG_FOURN'));
        // Element HT
        XPorcHt := CalculPort (True,TOBPortLOC);
        Base := GetBaseCoTrait (TOBasesLOC,TOBRG.GetString('PRG_FOURN'),true);
        ValInterm := arrondi ((Base - XporcHt)*(TauxRg /100),DEV.decimale);
        TOBRG.putvalue('PRG_MTHTRGDEV',ValInterm);
        // Element TTC
        XPorcTTC := CalculPort (False,TOBPortLOC);
        Base := GetBaseCoTrait (TOBasesLOC,TOBRG.GetString('PRG_FOURN'),false);
        ValInterm := arrondi ((bASE - XPorcTTC)*(TauxRg /100),DEV.decimale);
        TOBRG.putvalue('PRG_MTTTCRGDEV',ValInterm);

        IndiceRg := TOBRG.GetValue('INDICERG');
        // Element bases de tva
        CalculeTvaRg (TOBRG,TOBasesLOC,TOBBasesRGLoc,TOBPortLOC,DEv,TauxRg/100,IndiceRG,MontantTaxes);
        // Reajustement Ht ou TTC en fonction du type de retenue
        if TypeRg = 'HT' then
        begin
          TOBRG.putvalue('PRG_MTTTCRGDEV',TOBRG.GetValue('PRG_MTHTRGDEV')  +MontantTaxes);
        end else
        begin
          TOBRG.putvalue('PRG_MTHTRGDEV',TOBRG.GetValue('PRG_MTTTCRGDEV') -MontantTaxes);
        end;
        //
        ReAjusteMontants (TOBRG,DEV);
        //
        // Element Caution Banquaire
        if (TOBRG.getValue('PRG_NUMCAUTION')='') and (TOBRG.getValue('PRG_BANQUECP')<>'')then
        begin
          if TOBPiece.getValue('GP_NATUREPIECEG')='DBT' then
          begin
            ValInterm := arrondi ((TOBPIece.getValue('GP_TOTALTTC') - XporcTTC )*(TauxRg /100),V_PGI.okdecv);
            TOBRG.putvalue('PRG_CAUTIONMT',ValInterm);
            ValInterm := arrondi ((TOBPIece.getValue('GP_TOTALTTCDEV') - XporcTTC) *(TauxRg /100),DEV.decimale);
            TOBRG.putvalue('PRG_CAUTIONMTDEV',ValInterm);
          end;
        end;
        // passage des baseRG provisoire dans la TOB finale
        if TOBBasesRGLoc.detail.count > 0 then
        begin
          repeat
            TOBBasesRGLoc.detail[IInd].ChangeParent(TOBBasesRG,-1);
          Until TOBBasesRGLoc.detail.count =0;
        end;
      end;
    end;
  FINALLY
    TOBPorcsCot.free;
    TOBBasesCot.free;
    TOBBasesRGLoc.free;
  END;
end;

// bouttade
procedure CalculeRGSimple (TOBPorcs,TOBPIECE,TOBRG,TOBBASES,TOBBASESRG: TOB;DEV:Rdevise);

var ValInterm,TauxRg,MontantTaxes : double;
    IndiceRg : integer;
    TypeRg : string;
    XPorcHT,XPorcTTC : double;
begin
if TOBRG = nil then exit;
if TOBRG.GetValue('PRG_TAUXRG') <> 0 then
   begin
   // Taux retenue
   TAuxRg := TOBRG.GetValue('PRG_TAUXRG');
   TypeRg := TOBRG.GetValue('PRG_TYPERG');
   XPorcHt := 0;XPorcTTC := 0;
   if not TOBRG.getBoolean('PRG_MTMANUEL') then
   begin
   // Element HT
     //XPorcHt := CalculPort (True,TOBPorcs);
     //
   ValInterm := arrondi ((TOBPIece.getValue('GP_TOTALHTDEV') - XporcHt)*(TauxRg /100),DEV.decimale);
   TOBRG.putvalue('PRG_MTHTRGDEV',ValInterm);
   // Element TTC
   XPorcTTC := CalculPort (False,TOBPorcs);
   ValInterm := arrondi ((TOBPIece.getValue('GP_TOTALTTCDEV') - XPorcTTC)*(TauxRg /100),DEV.decimale);
   TOBRG.putvalue('PRG_MTTTCRGDEV',ValInterm);
     //
   end;
   IndiceRg := TOBRG.GetValue('INDICERG');
   // Element bases de tva
   CalculeTvaRg (TOBRG,TOBBases,TOBBasesRg,TOBPorcs,DEv,TauxRg/100,IndiceRG,MontantTaxes);
   // Reajustement Ht ou TTC en fonction du type de retenue
   if TypeRg = 'HT' then
      begin
      TOBRG.putvalue('PRG_MTTTCRGDEV',TOBRG.GetValue('PRG_MTHTRGDEV')  +MontantTaxes);
      end else
      begin
      TOBRG.putvalue('PRG_MTHTRGDEV',TOBRG.GetValue('PRG_MTTTCRGDEV') -MontantTaxes);
      end;
//
   ReAjusteMontants (TOBRG,DEV);
//
   // Element Caution Banquaire
   if (TOBRG.getValue('PRG_NUMCAUTION')='') and (TOBRG.getValue('PRG_BANQUECP')<>'')then
      begin
      if TOBPiece.getValue('GP_NATUREPIECEG')='DBT' then
      	begin
        ValInterm := arrondi ((TOBPIece.getValue('GP_TOTALTTC') - XporcTTC )*(TauxRg /100),V_PGI.okdecv);
        TOBRG.putvalue('PRG_CAUTIONMT',ValInterm);
        ValInterm := arrondi ((TOBPIece.getValue('GP_TOTALTTCDEV') - XporcTTC) *(TauxRg /100),DEV.decimale);
        TOBRG.putvalue('PRG_CAUTIONMTDEV',ValInterm);
      	end;
      end;
   end;

end;

function RecupDebutRG (TOBpiece : TOB;Indice : Integer): Integer;
var
   TOBL : TOB;
   Ind,ITypA: Integer;
begin
result := 0;
if Indice <= 0 then exit;
ITypA :=TOBPiece.detail[Indice].GetNumChamp ('GL_TYPEARTICLE');
for ind := Indice downto 0 do
    begin
    TOBL := TOBPiece.detail[Ind];
    if (TOBL.GetValeur (ITypA) = 'EPO') then
      begin
      // On prend en compte jusqu'a la retenue de garantie précédente
      result := Ind + 1;
      break;
      end;
    end;
end;

procedure VidePieceRGLigne (TOBPieceRG: TOB;Ligne : integer);
var TOBL : TOB;
begin
if TOBPieceRG = nil then exit;
TOBL := TOBPieceRG.findfirst (['INDICERG'],[Ligne],true);
while TOBL <> nil do
   begin
   TOBL.Putvalue('PRG_MTHTRG',0);
   TOBL.Putvalue('PRG_MTHTRGDEV',0);
   TOBL.Putvalue('PRG_MTTTCRG',0);
   TOBL.Putvalue('PRG_MTTTCRGDEV',0);
   if TOBL.getValue('PRG_NUMCAUTION') = '' then
      begin
      TOBL.Putvalue('PRG_CAUTIONMT',0);
      TOBL.Putvalue('PRG_CAUTIONMTDEV',0);
      end;
   TOBL := TOBPieceRG.findnext (['INDICERG'],[Ligne],true);
   end;
end;

procedure AlimenteValoLigneRgCot (TOBPiece,TOBLIG,TOBPieceRG,TOBBasesRG : TOB;EnHt:boolean; TypeL:RGmode);
var XP,XD,TP,TD : double;
begin
  if TypeL = RgStd then
  begin
    TOBLIG.PutValue('GL_QTEFACT',TOBPIECERG.detail[0].getValue('PRG_TAUXRG'));
    TOBLIG.PutValue('GL_PRIXPOURQTE',100);
    TOBLIG.PutValue('GL_QUALIFQTEVTE','%');
  end;
  GetValueRgCot (TOBPIECERG,XP,XD,True);
  if TOBBasesRG<>nil then
  begin
  	GetCumulTaxeByIndiceCot (TOBPIECE,TOBPIECERG,TOBBasesRG,TP,TD);
  end;
  if ((EnHt) and (TOBPIecerg.GetValue('PRG_TYPERG')='HT')) or
  ( (not EnHt) and (TOBPIecerg.GetValue('PRG_TYPERG')='TTC')) then
  begin
    if TypeL = RGstd then
    begin
      TOBLIG.PutValue('GL_PUHT',TOBPIECE.getValue('GP_TOTALHT'));
      TOBLIG.PutValue('GL_PUHTDEV',TOBPIECE.getValue('GP_TOTALHTDEV'));
      TOBLIG.PutValue('GL_PUTTC',TOBPIECE.getValue('GP_TOTALTTC'));
      TOBLIG.PutValue('GL_PUTTCDEV',TOBPIECE.getValue('GP_TOTALTTCDEV'));
    end;
    TOBLIG.PutValue('GL_MONTANTHT',XP-TP);
    TOBLIG.PutValue('GL_MONTANTHTDEV',XD-TD);
    TOBLIG.PutValue('GL_MONTANTTTC',XP);
    TOBLIG.PutValue('GL_MONTANTTTCDEV',XD);
  end else
  begin
  if TypeL = RGstd then
  begin
    TOBLIG.PutValue('GL_PUHT',TOBPIECE.getValue('GP_TOTALTTC'));
    TOBLIG.PutValue('GL_PUHTDEV',TOBPIECE.getValue('GP_TOTALTTCDEV'));
    TOBLIG.PutValue('GL_PUTTC',TOBPIECE.getValue('GP_TOTALHT'));
    TOBLIG.PutValue('GL_PUTTCDEV',TOBPIECE.getValue('GP_TOTALHTDEV'));
  end;
  TOBLIG.PutValue('GL_MONTANTHT',XP);
  TOBLIG.PutValue('GL_MONTANTHTDEV',XD);
  TOBLIG.PutValue('GL_MONTANTTTC',XP-TD);
  TOBLIG.PutValue('GL_MONTANTTTCDEV',XD-TD);
  end;
end;


procedure AlimenteValoLigneRg (TOBPiece,TOBLIG,TOBPieceRG,TOBBasesRG : TOB;EnHt:boolean; TypeL:RGmode);
var XP,XD,TP,TD : double;
begin
  if TypeL = RgStd then
   begin
   TOBLIG.PutValue('GL_QTEFACT',TOBPIECERG.getValue('PRG_TAUXRG'));
   TOBLIG.PutValue('GL_PRIXPOURQTE',100);
   TOBLIG.PutValue('GL_QUALIFQTEVTE','%');
   end;
  GetValueRg (TOBPIECERG,XP,XD,True);
  if TOBBasesRG<>nil then
   begin
   GetCumulTaxeByIndice (TOBPIECE,TOBPIECERG,TOBBasesRG,TP,TD);
   end;
  if ((EnHt) and (TOBPIecerg.GetValue('PRG_TYPERG')='HT')) or
  ( (not EnHt) and (TOBPIecerg.GetValue('PRG_TYPERG')='TTC')) then
  begin
  if TypeL = RGstd then
     begin
     TOBLIG.PutValue('GL_PUHT',TOBPIECE.getValue('GP_TOTALHT'));
     TOBLIG.PutValue('GL_PUHTDEV',TOBPIECE.getValue('GP_TOTALHTDEV'));
     TOBLIG.PutValue('GL_PUTTC',TOBPIECE.getValue('GP_TOTALTTC'));
     TOBLIG.PutValue('GL_PUTTCDEV',TOBPIECE.getValue('GP_TOTALTTCDEV'));
     end;
  TOBLIG.PutValue('GL_MONTANTHT',XP-TP);
  TOBLIG.PutValue('GL_MONTANTHTDEV',XD-TD);
  TOBLIG.PutValue('GL_MONTANTTTC',XP);
  TOBLIG.PutValue('GL_MONTANTTTCDEV',XD);
  end else
  begin
  if TypeL = RGstd then
     begin
     TOBLIG.PutValue('GL_PUHT',TOBPIECE.getValue('GP_TOTALTTC'));
     TOBLIG.PutValue('GL_PUHTDEV',TOBPIECE.getValue('GP_TOTALTTCDEV'));
     TOBLIG.PutValue('GL_PUTTC',TOBPIECE.getValue('GP_TOTALHT'));
     TOBLIG.PutValue('GL_PUTTCDEV',TOBPIECE.getValue('GP_TOTALHTDEV'));
     end;
  TOBLIG.PutValue('GL_MONTANTHT',XP);
  TOBLIG.PutValue('GL_MONTANTHTDEV',XD);
  TOBLIG.PutValue('GL_MONTANTTTC',XP-TD);
  TOBLIG.PutValue('GL_MONTANTTTCDEV',XD-TD);
  end;
end;
//
procedure RGCotVersLigne (TOBPiece,TOBPieceRG,TOBBasesRG,TOBLIG,TOBLigneRG : TOB; Applicretenue: boolean; GenerationFac : boolean=false);
var EnHt:boolean;
    CautionUsed , CautionUsedDev,MtRG,MyResiduRg : double;
    MyCautionRest : double;
    MaxNumOrdre,Indice : integer;
    Intervenant : string;
begin
  CautionUsed := 0; CautionUsedDev := 0;
  if TOBPieceRG = nil then exit;
  EnHt:=(TOBPiece.GetValue('GP_FACTUREHT')='X');
  TOBLIG.Putvalue('GL_TYPELIGNE','RG');
  MtRg := 0;
  for Indice := 0 to TOBPieceRG.detail.count -1 do
  begin
  	CautionUsed := 0;
    CautionUsedDev := 0;
  	MyCautionRest:=0;
    if (GenerationFac) and (TOBPieceRG.detail[Indice].fieldExists ('PIECEPRECEDENTE')) and
      ( Pos(TOBPieceRG.detail[Indice].getValue('PRG_NATUREPIECEG'),'ABT;ABP')=0) then
    begin
      Intervenant := TOBPieceRG.detail[Indice].GetString('PRG_FOURN');
      GetCautionAlreadyUsed (TOBPiece,TOBLIG,TOBPieceRG.detail[Indice].GetString('PIECEPRECEDENTE'),CautionUsed,CautionUsedDev,Intervenant);
    end;
    if TOBPIeceRG.detail[Indice].GetValue('PRG_CAUTIONMTDEV') <> 0 then
    begin
      MyCautionRest := TOBPIeceRG.detail[Indice].GetValue('PRG_CAUTIONMTDEV') - CautionUsedDev ;
    end;
    if MyCautionRest < 0 then MyCautionrest := 0;
    MyResiduRg := TOBPIeceRG.detail[Indice].GetDouble('PRG_MTTTCRGDEV') - MyCautionrest;
    if MyResiduRg < 0 then MyResiduRg := 0;
    MtRg := MtRg + MyResiduRg;
  end;
  if MtRg < 0 then MtRg := 0;
  if (TOBPiece.getValue('GP_NATUREPIECEG')<>'ABT') and (TOBPiece.getValue('GP_NATUREPIECEG')<>'ABP') then
  begin
    if MtRg = 0 then
    begin
    //Caution Bancaire
      TOBLIG.Putvalue('GL_LIBELLE','Caution Bancaire sur Retenue de Garantie')
    end else
    begin
      AlimenteValoLigneRgCot (TOBPiece,TOBLIG,TOBPieceRG,TOBBasesRG,EnHt,RgDepass);
      TOBLIG.Putvalue('GL_LIBELLE','Retenue de Garantie '+TOBPIeceRg.detail[0].getValue('PRG_TYPERG'));
      TOBLIG.putValue('GL_MONTANTHTDEV',MtRg);
      //
    end;
  end;
  if not ApplicRetenue then TOBLIG.Putvalue('GL_NONIMPRIMABLE','X');
  if (TOBLigneRG <> nil) and (TOBLigneRG.detail[0].FieldExists('PIECEPRECEDENTE')) then TOBLIG.PutVAlue ('GL_PIECEPRECEDENTE',TOBLigneRG.detail[0].GetValue('PIECEPRECEDENTE'));
  if (TOBPieceRG <> nil) and (TOBPieceRG.detail[0].FieldExists('NUMORDRE')) and (TOBPieceRG.detail[0].GEtVAlue('NUMORDRE') <> 0)  then
  begin
  	TOBLIG.PutValue('GL_NUMORDRE',TOBPieceRG.detail[0].GEtVAlue('NUMORDRE'));
  end else
  begin
    MaxNumOrdre := GetMaxNumOrdre (TOBPiece);
    TOBLIG.PutValue('GL_NUMORDRE',MaxNumordre);
  end;

end;

//
procedure RGMulVersLigne (TOBPiece,TOBPieceRG,TOBLIG,TOBLigneRG : TOB; Applicretenue: boolean; GenerationFac : boolean=false);
var Xp,XD,XPTTC,XDTTC : double;
		NumCaution,TypeRg : string;
    TauxRg : double;
    MaxNumOrdre : integer;
    EnHt : boolean;
begin
  if TOBPieceRG.detail.Count = 0 then exit;
  TypeRg := TOBPIeceRg.detail[0].getString('PRG_TYPERG');
  TauxRg := TOBPIeceRg.detail[0].getDouble('PRG_TAUXRG');
  GetRg (TOBpieceRG,true,True,Xp,XD,NumCaution);
  GetRg (TOBpieceRG,false,True,XpTTC,XDTTC,NumCaution);
  EnHt := (TypeRg='HT');
  TOBLIG.Putvalue('GL_TYPELIGNE','RG');
  if NumCaution <> '' then
  begin
  	TOBLIG.Putvalue('GL_LIBELLE','Caution Bancaire sur Retenue de Garantie')
  end else
  begin

  	TOBLIG.Putvalue('GL_LIBELLE','Retenue de Garantie '+TypeRG+' de '+floattostr(TauxRg)+ '%');
    TOBLIG.PutValue('GL_QTEFACT',TauxRg);
    TOBLIG.PutValue('GL_PRIXPOURQTE',100);
    TOBLIG.PutValue('GL_QUALIFQTEVTE','%');

    if ((EnHt) and (TOBPIecerg.GetValue('PRG_TYPERG')='HT')) or
    ( (not EnHt) and (TOBPIecerg.GetValue('PRG_TYPERG')='TTC')) then
    begin
      TOBLIG.PutValue('GL_PUHT',TOBPIECE.getValue('GP_TOTALHT'));
      TOBLIG.PutValue('GL_PUHTDEV',TOBPIECE.getValue('GP_TOTALHTDEV'));
      TOBLIG.PutValue('GL_PUTTC',TOBPIECE.getValue('GP_TOTALTTC'));
      TOBLIG.PutValue('GL_PUTTCDEV',TOBPIECE.getValue('GP_TOTALTTCDEV'));
    end else
    begin
      TOBLIG.PutValue('GL_PUHT',TOBPIECE.getValue('GP_TOTALTTC'));
      TOBLIG.PutValue('GL_PUHTDEV',TOBPIECE.getValue('GP_TOTALTTCDEV'));
      TOBLIG.PutValue('GL_PUTTC',TOBPIECE.getValue('GP_TOTALHT'));
      TOBLIG.PutValue('GL_PUTTCDEV',TOBPIECE.getValue('GP_TOTALHTDEV'));
    end;

    TOBLIG.PutValue('GL_MONTANTHT',XP);
    TOBLIG.PutValue('GL_MONTANTHTDEV',XD);
    TOBLIG.PutValue('GL_MONTANTTTC',XPTTC);
    TOBLIG.PutValue('GL_MONTANTTTCDEV',XDTTC);
  end;
  if not ApplicRetenue then TOBLIG.Putvalue('GL_NONIMPRIMABLE','X');
  (*
  if (TOBLigneRG <> nil) and (TOBLigneRG.FieldExists('PIECEPRECEDENTE')) then TOBLIG.PutVAlue ('GL_PIECEPRECEDENTE',TOBLigneRG.GetValue('PIECEPRECEDENTE'));
  if (TOBPieceRG <> nil) and (TOBPieceRG.FieldExists('NUMORDRE')) and (TOBPieceRG.GEtVAlue('NUMORDRE') <> 0)  then
  begin
  	TOBLIG.PutValue('GL_NUMORDRE',TOBPieceRG.GEtVAlue('NUMORDRE'));
  end else
  begin
    MaxNumOrdre := GetMaxNumOrdre (TOBPiece);
    TOBLIG.PutValue('GL_NUMORDRE',MaxNumordre);
  end;
  *)
end;


procedure RGVersLigne (TOBPiece,TOBPieceRG,TOBBasesRG,TOBLIG,TOBLigneRG : TOB; Applicretenue: boolean; GenerationFac : boolean=false);
var EnHt:boolean;
    CautionUsed , CautionUsedDev : double;
    MyCautionRest : double;
    MaxNumOrdre : integer;
begin
  if TOBPieceRG = nil then exit;
  if TOBPieceRG.detail.count > 0 then
  begin
    RGMulVersLigne (TOBpiece,TOBPieceRG,TOBLIG,TOBLigneRG,Applicretenue,GenerationFac);
    exit;
  end;
  CautionUsed := 0; CautionUsedDev := 0;
  if (GenerationFac) and (TOBPieceRG.fieldExists ('PIECEPRECEDENTE')) and ( Pos(TOBPieceRG.getValue('PRG_NATUREPIECEG'),'ABT;ABP')=0) then
  begin
	  GetCautionAlreadyUsed (TOBPiece,TOBLIG,TOBPieceRG.getVAlue('PIECEPRECEDENTE'),CautionUsed,CautionUsedDev);
  end;
  MyCautionRest:=0;
  EnHt:=(TOBPiece.GetValue('GP_FACTUREHT')='X');
  TOBLIG.Putvalue('GL_TYPELIGNE','RG');
  if TOBPIeceRG.GetValue('PRG_CAUTIONMTDEV') <> 0 then
  begin
    MyCautionrest := TOBPIeceRG.GetValue('PRG_CAUTIONMTDEV') - CautionUsedDev;
    if MyCautionRest < 0 then MyCautionRest := 0;
  end;
  if (TOBPIECERG.GetValue('PRG_NUMCAUTION')<> '') and ( Pos(TOBPiece.getValue('PRG_NATUREPIECEG'),'ABT;ABP')=0) then
  begin
    if TOBPIECERG.GetValue('PRG_MTTTCRGDEV') <= (MyCautionrest) then
    begin
    //Caution Bancaire
      TOBLIG.Putvalue('GL_LIBELLE','Caution Bancaire sur Retenue de Garantie')
    end else
    begin
      AlimenteValoLigneRg (TOBPiece,TOBLIG,TOBPieceRG,TOBBasesRG,EnHt,RgDepass);
      TOBLIG.Putvalue('GL_LIBELLE','RG sur dépassement de caution '+TOBPIeceRg.getValue('PRG_TYPERG'));
      //
      TOBLIG.putValue('GL_MONTANTHTDEV',Abs(MyCautionrest - TOBPIeceRG.GetValue('PRG_MTTTCRGDEV')));
      //
    end;
  end else
  begin
    AlimenteValoLigneRg (TOBPiece,TOBLIG,TOBPieceRG,TOBBasesRG,EnHt,RGstd);
    TOBLIG.Putvalue('GL_LIBELLE','Retenue de Garantie '+TOBPIeceRg.getValue('PRG_TYPERG')+' de '+floattostr(TOBPIECERG.GetValue('PRG_TAUXRG'))+ '%');
  end;
  if not ApplicRetenue then TOBLIG.Putvalue('GL_NONIMPRIMABLE','X');
  if (TOBLigneRG <> nil) and (TOBLigneRG.FieldExists('PIECEPRECEDENTE')) then TOBLIG.PutVAlue ('GL_PIECEPRECEDENTE',TOBLigneRG.GetValue('PIECEPRECEDENTE'));
  if (TOBPieceRG <> nil) and (TOBPieceRG.FieldExists('NUMORDRE')) and (TOBPieceRG.GEtVAlue('NUMORDRE') <> 0)  then
  begin
  	TOBLIG.PutValue('GL_NUMORDRE',TOBPieceRG.GEtVAlue('NUMORDRE'));
  end else
  begin
    MaxNumOrdre := GetMaxNumOrdre (TOBPiece);
    TOBLIG.PutValue('GL_NUMORDRE',MaxNumordre);
  end;

end;

procedure CalculBaseRG (TOBTiers,TOBPiece,TOBPieceTrait,TOBBasesL,TOBPIeceRG,TOBBasesRg,TOBS,TOBPorcs,TOBBASESG: TOB; R_Valeurs : R_CPercent;DEV:Rdevise;GenerationFac:boolean=false);
var
   TOBL,TOBBases,TOBPIECELOC,TOBRG,TOBLLOC,TOBPorcsLoc,TOBBasesLoc,TOBBL,TOBPieceRGLoc,TOBBasesRGLoc: TOB;
   Indice,INdL,IndiceRg : Integer;
   IPuHt,IpuHtdev,IpuTTC,iputtcdev,iMontHt,imontHtDev,iMontTTC : integer;
   iMontTTCdev: integer;
   bid1,bid2,bid3,bid4 : double;
   Applicretenue : boolean;
   PortHt,PortTTC : double;
   NumOrdre,IInd : integer;
   TOBB : TOB;
   TOBTaxes : TOB;
begin
  if TOBPieceRG = nil then exit;
  ApplicRetenue:=(GetInfoParPiece(TOBPIECE.GetValue('GP_NATUREPIECEG'),'GPP_APPLICRG')='X');
  TOBPIECELOC:=TOB.create ('PIECE',nil,-1);
  TOBPorcsLoc := TOB.Create ('LES PORTS LOC',nil,-1);
  TOBBasesLoc := TOB.Create ('LES BASES LIGNES',nil,-1);
  TOBLLOC:=TOB.Create('LIGNE',nil,-1);
  TOBBases := TOB.create('BASES PIED(LIGNES)',nil,-1);
  TOBPieceRGLoc := TOB.create ('LES RG',nil,-1);
  TOBBasesRGLoc := TOB.create ('LES BASES RG',nil,-1);
  //
  AddLesSupEntete (TOBpieceLoc);
  TOBPieceLOC.dupliquer (TOBPIECE,false,true);
  ZeroFacture (TOBpieceloc);
  //
  AddLesSupLigne (TOBLLOc,false);
  //
  IndiceRg := TOBS.GetValue('INDICERG');
  TOBPorcsLoc.Dupliquer (TOBPorcs,true,true);
  //
  TOBBasesLoc.Dupliquer (TOBBasesL,true,true);
  // nettoyage avant création
  VidePieceRGLigne (TOBPieceRG,IndiceRG);
  // --
  ZeroLignePourcent(TOBS) ;
  ZeroLignePourcent(TOBLLOC) ;
  IndL := 0;
  IPuHt := 0;
  IPuHtdev := 0;
  IPuTTC := 0;
  IPuttcdev := 0;
  IMOntHt := 0;
  IMOntHtDev := 0;
  IMOntttc := 0;
  IMOntttcDev := 0;
  for Indice := R_Valeurs.Depart to R_Valeurs.Fin do
  begin
    TOBL := GetTOBLigne(TOBPiece,Indice+1 ) ; if TOBL = nil then break;
    NumOrdre := TOBL.getValue('GL_NUMORDRE');
    if Indice = R_Valeurs.Depart then
    begin
      IndL := TOBL.GetNumChamp ('GL_TYPELIGNE');
      // PU HT
      IPuHtDev := TOBL.GetNumChamp ('GL_PUHTDEV');
      IPuHt := TOBL.GetNumChamp ('GL_PUHT');
      // PU TTC
      IPuTTCDev := TOBL.GetNumChamp ('GL_PUTTCDEV');
      IPuTTC := TOBL.GetNumChamp ('GL_PUTTC');
      // Montant Ht
      IMontHt := TOBL.GetNumChamp ('GL_TOTALHT');
      IMontHtDev := TOBL.GetNumChamp ('GL_TOTALHTDEV');
      // Montant TTC
      IMontTTC := TOBL.GetNumChamp ('GL_TOTALTTC');
      IMontTTCDev := TOBL.GetNumChamp ('GL_TOTALTTCDEV');
      // -------
    end;
    if TOBL.GetValeur(INDL) = 'EPO' then break;
    // VARIANTE
    //    if isVariante(TOBL) then continue;
    if (isVariante(TOBL))
    {$IFDEF BTP}
    		or (IsCentralisateurBesoin(TOBL))
		    or (TOBL.GetValue('GL_TYPEARTICLE') = '$#$') then continue;
    {$ENDIF}
    // --
    if (TOBL.getvaleur(INDL) = 'RG' ) then continue;
    if (TOBL.getvalue('GL_ARTICLE') = '' ) then continue;
    SommeLignePiedHT (TOBL,TOBpieceLoc);
    TOBBL := FindLignesbases (numordre,TOBBasesL);
    if TOBBL <> nil then CumuleLesBases (TOBBases, TOBBL,DEV.Decimale);
  end;
  // --
  RecalculPiedPort (tamodif,TOBpieceLoc,TOBporcsloc,TOBBASESG);
  AddlesPorts (TOBPIECELOC,TOBPorcsLoc,TOBBases,TOBBasesLoc,TOBtiers,DEv,TaCreat);
  if TOBpieceRG.detail.count > 0 then
  begin
  indice := 0;
  repeat
    TOBRG := TOBpieceRG.detail[Indice];
    if TOBRG.getValue('INDICERG') <> IndiceRg then
    begin
      inc(indice);
    	continue;
    end;
    TOBRG.ChangeParent (TOBPieceRGLoc,-1);
  until indice >= TOBpieceRG.detail.count ;
  end;
  if TOBBasesRG.detail.count > 0 then
  begin
    indice := 0;
    repeat
      TOBRG := TOBbasesRG.detail[Indice];
      if TOBRG.getValue('INDICERG') <> IndiceRg then
      begin
        inc(indice);
        continue;
      end;
      TOBRG.ChangeParent (TOBBasesRGLoc,-1);
    until indice >= TOBbasesRG.detail.count ;
  end;
  if TOBPieceRGLoc <> nil then
  begin
		RecalculeRG (TOBPorcsLoc,TOBPIECELOC,TOBPieceRGLoc,TOBBases,TOBBasesRGLoc,TOBPieceTrait,DEV);
//    CalculeRGSimple (TOBPorcsLoc,TOBPIECELOC,TOBRG,TOBBases,TOBBASESRG,DEV);
		RGVersLigne (TOBPIECELOC,TOBPieceRGLoc,TOBBasesRGLoc,TOBS,nil,Applicretenue,GenerationFac);
    IInd := 0;
    if TOBPieceRGLoc.detail.count > 0 then
    begin
      repeat
        TOBRG := TOBPieceRGLoc.detail[IInd];
        TOBRG.ChangeParent(TOBPieceRG,-1);
      until TOBPieceRGLoc.detail.count = 0;
    end;
    IInd := 0;
    if TOBBasesRGloc.detail.count > 0 then
    begin
      repeat
        TOBRG := TOBbasesRGLoc.detail[IInd];
        TOBRG.ChangeParent(TOBBasesRg,-1);
      until TOBbasesRGLoc.detail.count = 0;
    end;
  end;
  // nettoyage avant départ
  TOBPIECELOC.free;
  TOBPorcsLoc.free;
  TOBBasesLoc.free;
  TOBLLOC.free;
  TOBBases.free;
  TOBPieceRGLoc.free;
  TOBbasesRGLoc.free;
end;

function RGMultiple (TOBpiece:TOB):boolean;
var nbr:integer;
		TOBLL : TOB;
begin
  result := false;
  if TOBpiece.getValue('GP_RGMULTIPLE')='X' then
  begin
  	result := true;
    exit;
  end;
end;

procedure InsereLigneRGCot(TOBPiece,TOBPieceRG,TOBBASEs,TOBBRG,TOBTIers,TOBAffaire,TOBLigneRG: TOB;NumIndice:integer);
var TOBLoc,TOBBRGL:TOB;
    numl,i : integer;
    ApplicRetenue: boolean;
begin
  if TOBPieceRG = nil then exit;
  if RGMultiple(TOBpiece) then exit;

  ApplicRetenue:=(GetInfoParPiece(TOBPIECE.GetValue('GP_NATUREPIECEG'),'GPP_APPLICRG')='X');
  if TOBPieceRG.detail.count = 0 then exit;
  if TOBPieceRG.detail[0].getValue('PRG_TAUXRG') = 0 then exit;
  TOBLOC:=NewTOBLigne (TOBPIECE,numIndice+1);
  if NumIndice = -1 then NumL:=TOBPiece.Detail.Count else NumL:=NumIndice+1;
  InitialiseTOBLigne(TOBPiece,TOBTiers,TOBAffaire,NumL) ;
  RGCotVersLigne (TOBPiece,TOBPieceRG,TOBBRG,TOBLOC,TobLigneRG,ApplicRetenue);
  TOBLOC.putvalue('INDICERG',1);
  //
  for i:=0 to TOBPieceRg.detail.count-1 do
  begin
    TOBBRGL := TOBPieceRG.detail[i];
    TOBBRGL.putvalue('PRG_NUMLIGNE',Numl);
    TOBBRGL.putvalue('INDICERG',1);
  end;
  for i:=0 to TOBBRG.detail.count-1 do
  begin
    TOBBRGL := TOBBRG.detail[i];
    TOBBRGL.putvalue('PBR_NUMLIGNE',Numl);
    TOBBRGL.putvalue('INDICERG',1);
  end;
end;


procedure InsereLigneRG(TOBPiece,TOBPieceRg,TOBBASEs,TOBBRG,TOBTIers,TOBAffaire,TOBLigneRG: TOB;NumIndice:integer);
var TOBLoc,TOBBRGL:TOB;
    numl,i : integer;
    ApplicRetenue: boolean;
    TOBRg : TOB;
begin
  if TOBPieceRg = nil then exit;
  if TOBPieceRg.detail.count =0 then exit;
  if RGMultiple(TOBpiece) then exit;
  ApplicRetenue:=(GetInfoParPiece(TOBPIECE.GetValue('GP_NATUREPIECEG'),'GPP_APPLICRG')='X');
  TOBRg := TOBPieceRG.detail[0];
  if TOBRG.getValue('PRG_TAUXRG') = 0 then exit;
  TOBLOC:=NewTOBLigne (TOBPIECE,numIndice+1);
  if NumIndice = -1 then NumL:=TOBPiece.Detail.Count else NumL:=NumIndice+1;
  InitialiseTOBLigne(TOBPiece,TOBTiers,TOBAffaire,NumL) ;
  RGVersLigne (TOBPiece,TOBRG,TOBBRG,TOBLOC,TobLigneRG,ApplicRetenue);
  TOBLOC.putvalue('INDICERG',1);
  TOBRG.putvalue('PRG_NUMLIGNE',Numl);
  TOBRG.putvalue('INDICERG',1);
  for i:=0 to TOBBRG.detail.count-1 do
  begin
    TOBBRGL := TOBBRG.detail[i];
    TOBBRGL.putvalue('PBR_NUMLIGNE',Numl);
    TOBBRGL.putvalue('INDICERG',1);
  end;
end;

function IsRGSimple (TOBPieceRG : TOB) : boolean;
var indice : integer;
		LastNum : integer;
begin
  result := true;
  for Indice := 0 to TOBPieceRG.detail.count -1 do
  begin
    if Indice = 0 then
    begin
  		LastNum := TOBPieceRG.detail[0].getValue('PRG_NUMLIGNE');
    end else
    begin
    	if LastNum <> TOBPieceRG.detail[indice].getValue('PRG_NUMLIGNE') then
      begin
        result := false;
        break;
      end;
    end;
  end;
end;

procedure RecalculeRG (TOBPorcs,TOBPIECE,TOBPIECERG,TOBBASES,TOBBASESRG,TOBPieceTrait : TOB;DEV:Rdevise);
var TOBRG : TOB;
BEGIN
  if TOBPIECERG = nil then exit;
  (*
  if ((TOBPieceRG.detail.count > 1) and (IsRGSimple(TOBPieceRG))) or (TOBPieceTRait.detail.count > 1) then
  begin
    CalculeRGSimpleCotrait (TOBPorcs,TOBPIECE,TOBPieceRG,TOBBASES,TOBBASESRG,TOBPieceTrait,DEV);
  end else if TOBPieceRG.detail.count = 1 then
  begin
  *)
  if TOBPieceRg.detail.count > 0 then
  begin
    TOBRG := TOBPIECERG.detail[0];
    CalculeRGSimple (TOBPorcs,TOBPIECE,TOBRG,TOBBASES,TOBBASESRG,DEV);
  end;
//  end;
END;

function GetCumultextRG (TOBPIeceRG:TOB;Nomchamp : string) : string;
var ind: integer;
    TOBLOC : TOB;
begin
  result := '';
  if TOBPieceRG = nil then exit;
  for ind:=0 to TOBPIECERG.detail.count -1 do
  begin
    TOBLOc := TOBPIECERG.detail[ind];
    Result:= TOBLoc.GetValue(nomchamp);
  end;
end;

function GetCumulValueRG (TOBPIeceRG:TOB;Nomchamp : string) : variant;
var ind: integer;
    TOBLOC: TOB;
begin
result := 0;
if TOBPieceRG = nil then exit;
for ind:=0 to TOBPIECERG.detail.count -1 do
    begin
    TOBLOc := TOBPIECERG.detail[ind];
    if (Nomchamp='PRG_TAUXRG') then
       begin
       Result:= TOBLOC.GetValue(nomchamp);
       break;
       end;
    if (Nomchamp='PRG_MTHTRGDEV') or (Nomchamp='PRG_MTTTCRGDEV') then
       begin
       result := result + TOBLOC.GetValue(nomchamp);
       end;
    end;
end;

function GetCumulValueTaxeRG (TOBBasesRG:TOB;Nomchamp : string) : variant;
var ind: integer;
    TOBLOC: TOB;
begin
result := 0;
if TOBbasesRG = nil then exit;
for ind:=0 to TOBbasesRG.detail.count -1 do
    begin
    TOBLOc := TOBBasesRG.detail[ind];
    if (Nomchamp='PBR_VALEURDEV') then
       begin
       Result:= result+TOBLOC.GetValue(nomchamp);
       break;
       end;
    end;
end;

procedure GetMontantRG (TOBPIeceRG,TOBBasesRG : TOB;var XD,XP : double;DEV:Rdevise;HtForce:boolean=false;HorsCaution:boolean=false);
var XXD,XXP : double;
    XTD,XTP : double;
    Enht : Boolean;
begin
XD := 0; XP :=0;
if TOBPieceRG = nil then exit;
EnHt := (GetCumulTextRG (TOBPieceRG,'PRG_TYPERG') = 'HT');
GetCumulRG (TOBPieceRG,XXP,XXD,not HorsCaution ); // Valeur residuelle TTC de RG
XD := XXD;
XP := XXP;
if (EnHt) or (HtForce) then
   begin
   GetCumulTaxesRG (TOBBasesRG,TOBPieceRG,XTD,XTP,DEV); // Tva Residuelle
   XD := XXD - XTD;  // HT = TTC - TVA
   XP := XXP - XTP;  // HT = TTC - TVA
   end;
end;

procedure GetCumulRG (TOBPieceRG : TOB; var SommePivot,SommeDevise : double;initial:boolean=false);
var TOBRG : TOB;
    indice : integer;
    MyCautionRest : double;
    LocPivot,LocDevise,Sens,MtCaution,MtCautionU : double;
begin
  SommePivot:=0; SommeDevise := 0;
  if (TOBPieceRG=nil) then exit;
  for indice:=0 to TOBPieceRG.detail.count -1 do
  begin
    TOBRG := TOBPIeceRG.detail[Indice];
    if TOBRG.GetValue('PRG_MTTTCRG') < 0 then Sens := -1 else Sens := 1;
    LocPivot := Abs(TOBRG.GetValue('PRG_MTTTCRG'));
    LocDevise := Abs(TOBRG.GetValue('PRG_MTTTCRGDEV'));
    //
    if (TOBRG.GetValue('PRG_NUMCAUTION')<> '') and (not initial) then
    begin
//      MyCautionRest := TOBRG.GetValue('PRG_CAUTIONMT') - TOBRG.GetValue('PRG_CAUTIONMTU');
      MtCaution := Abs(TOBRG.GetValue('PRG_CAUTIONMT'));
      MtCautionU := Abs(TOBRG.GetValue('PRG_CAUTIONMTU'));
      MyCautionRest := MtCaution - MtCautionU;
      //
      if MyCautionRest < 0 then MyCautionRest := 0;
      LocPivot:= LocPivot - (MyCautionRest);
      //
//      MyCautionRest := TOBRG.GetValue('PRG_CAUTIONMTDEV') - TOBRG.GetValue('PRG_CAUTIONMTUDEV');
      MtCaution := Abs(TOBRG.GetValue('PRG_CAUTIONMTDEV'));
      MtCautionU := Abs(TOBRG.GetValue('PRG_CAUTIONMTUDEV'));
      MyCautionRest := MtCaution - MtCautionU;
      //
      if MyCautionRest < 0 then MyCautionRest := 0;
      LocDevise := LocDevise - (MyCautionRest);
      if (TOBRG.GetValue('PRG_NATUREPIECEG')<>'ABT') and (TOBRG.GetValue('PRG_NATUREPIECEG')<>'ABP') then
      begin
        if LocPivot < 0 then LocPivot:=0;
        if LocDevise < 0 then LocDevise:=0;
      end;
    end;
    //
    LocPivot := LocPivot * Sens;
    LocDevise := LocDevise * Sens;
    //
    SommePivot := SommePivot + LocPivot;
    SommeDevise := SommeDevise + LocDevise;
  end;
end;

procedure GetRGByIndice (TOBPieceRG:TOB; Indice : integer;var RGRP,RGRD:double; Reliquat: boolean=false);
var TOBLOC : TOB;
begin
RGRP:=0;RGRD:=0;
if TOBPieceRG=nil then exit;
TOBLOC:=TOBPIECERG.findfirst(['INDICERG'],[Indice],true);
if TOBLOC <> nil then
   begin
   GetValueRg (TOBLoc,RGRP,RGRD,Reliquat);
   end;
end;

procedure GetRGByLigne (TOBPieceRG:TOB; Ligne : integer;var RGRP,RGRD:double; Reliquat: boolean=false);
var TOBLOC : TOB;
begin
RGRP:=0;RGRD:=0;
if TOBPieceRG=nil then exit;
TOBLOC:=TOBPIECERG.findfirst(['PRG_NUMLIGNE'],[Ligne],true);
if TOBLOC <> nil then
   begin
   GetValueRg (TOBLoc,RGRP,RGRD,Reliquat);
   end;
end;

procedure RecupPosition (TOBBasesRG:TOB; CatTaxe,CodeTaxe:String; var Mode:integer; var Indice:integer);
var TOBB : TOB;
    ind : integer;
begin
Indice:=-1;
if TOBBasesRG = nil then exit;
Mode:= TRGModeInsertion;
for ind:=0 to TOBBasesRG.detail.count -1 do
    begin
    TOBB:=TOBBasesRG.detail[Ind];
    if (CatTAxe = TOBB.GetValue ('PBR_CATEGORIETAXE')) and (CodeTaxe = TOBB.GetValue ('PBR_FAMILLETAXE')) then
       begin
       Mode := TRGModeModif;
       Indice := ind;
       break;
       end;
    end;
end;

procedure InsereOuUpdate (TOBBasesRG,TOBBasesGlob,TOBPieceRG,TOBTvaMillieme : TOB);
var Indice : integer;
    Mode : integer;
    CatTaxe,Codetaxe : string;
    TOBLOC : TOB;
    DEV:Rdevise;
    Ratio,Millieme : double;
    RGRP,RGRD,RGD,RGP : double;
begin
if TOBBasesRg = nil then exit;
Millieme := 1;
CatTaxe := TOBBasesRg.GetValue('PBR_CATEGORIETAXE');
CodeTaxe := TOBBasesRg.GetValue('PBR_FAMILLETAXE');
(*
TOBMILLIEME := TOBTVAMillieme.findFirst(['BPM_CATEGORIETAXE','BPM_FAMILLETAXE'],[CatTaxe,CodeTaxe],true);
if TOBMillieme <> nil then Millieme := TOBMillieme.GetValue('BPM_MILLIEME')/1000;
*)
RecupPosition ( TOBBasesGlob,CatTaxe,CodeTaxe,Mode,Indice);
if Mode=TRGModeInsertion then
   begin
   TOBLOC:=TOB.create('PIEDBASERG',TOBBasesGlob,Indice);
   TOBLOC.InitValeurs;
   TOBLOC.addChampsupvaleur ('INDICERG',1);
   TOBLOC.putvalue('PBR_NATUREPIECEG',TOBBasesRG.GetValue('PBR_NATUREPIECEG'));
   TOBLOC.putvalue('PBR_DATEPIECE',TOBBasesRG.GetValue('PBR_DATEPIECE'));
   TOBLOC.putvalue('PBR_SOUCHE',TOBBasesRG.GetValue('PBR_SOUCHE'));
   TOBLOC.putvalue('PBR_NUMERO',TOBBasesRG.GetValue('PBR_NUMERO'));
   TOBLOC.putvalue('PBR_INDICEG',TOBBasesRG.GetValue('PBR_INDICEG'));
   TOBLOC.putvalue('PBR_CATEGORIETAXE',TOBBasesRG.GetValue('PBR_CATEGORIETAXE'));
   TOBLOC.putvalue('PBR_FAMILLETAXE',TOBBasesRG.GetValue('PBR_FAMILLETAXE'));
   TOBLOC.putvalue('PBR_TAUXDEV',TOBBasesRG.GetValue('PBR_TAUXDEV'));
   TOBLOC.putvalue('PBR_COTATION',TOBBasesRG.GetValue('PBR_COTATION'));
   TOBLOC.putvalue('PBR_DATETAUXDEV',TOBBasesRG.GetValue('PBR_DATETAUXDEV'));
   TOBLOC.putvalue('PBR_DEVISE',TOBBasesRG.GetValue('PBR_DEVISE'));
   TOBLOC.putvalue('PBR_SAISIECONTRE',TOBBasesRG.GetValue('PBR_SAISIECONTRE'));
   end else
   begin
   TOBLOC:=TOBBasesGlob.detail[Indice];
   end;
   DEV.Code:=TOBBasesRG.getValue('PBR_DEVISE') ; GetInfosDevise(DEV) ;
   // recupere le reliquat de RG par ligne
   GetRGByIndice (TOBPieceRG, TOBBasesRG.GetValue('INDICERG'),RGRP,RGRD,True);
   GetRGByIndice (TOBPieceRG, TOBBasesRG.GetValue('INDICERG'),RGP,RGD);
   if RGD <> 0 then Ratio := RGRD / RGD else Ratio := 0;
   //
//   TOBLOC.PutValue('PBR_BASETAXE',RGRP) ;
//   TOBLOC.PutValue('PBR_BASEDEV',RGRD) ;
   TOBLOC.PutValue('PBR_BASETAXE',TOBLOC.GetValue('PBR_BASETAXE')+(TOBBasesRG.GetValue('PBR_BASETAXE')*Ratio)) ;
   TOBLOC.PutValue('PBR_BASEDEV',TOBLOC.GetValue('PBR_BASEDEV')+(TOBBasesRG.GetValue('PBR_BASEDEV')*Ratio)) ;

   TOBLOC.PutValue('PBR_VALEURTAXE',TOBLOC.GetValue('PBR_VALEURTAXE') + (TOBBasesRg.GetValue('PBR_VALEURTAXE')*Ratio)) ;
   if DEV.code <>V_PGI.DevisePivot then
      TOBLOC.PutValue('PBR_VALEURDEV',TOBLOC.GetValue('PBR_VALEURDEV') + (TOBLOC.GetValue('PBR_VALEURDEV')*Ratio))
      else TOBLOC.PutValue('PBR_VALEURDEV',TOBLOC.GetValue('PBR_VALEURTAXE') );
end;

procedure FusionneTva (TOBTAXERG,TOBTAXEGLOB,TOBPIECERG,TOBTvaMillieme: TOB; Avoir : boolean=false);
var indice : integer;
    TOBLOc : TOB;
    DEV : Rdevise;
    XD,Xp,TheMontantRGTTC : double;
begin
TOBTaxeGlob.ClearDetail;
if TOBTaxeRg = nil then exit;
//
GetCumulRG (TOBPIeceRG,XP,XD,True);
TheMontantRGTTC := XD;
GetReliquatCaution (TOBPieceRg,XP,XD);
TheMontantRGTTC := TheMontantRGTTC - XD;
if not Avoir then
begin
  if TheMontantRGTTC < 0 then TheMontantRGTTC := 0;
end;
if TheMontantRgTTc = 0 then exit;
//

for Indice := 0 to TOBTaxeRG.detail.count -1 do
    begin
    TOBLOC:=TOBTaxeRG.detail[Indice];
    InsereouUpdate (TOBLOC,TOBTAXEGLOB,TOBPieceRG,TOBTvaMillieme);
    end;
for Indice := 0 to TOBTaxeGLOB.detail.count -1 do
    begin
    TOBLOC:=TOBTaxeGlob.detail[Indice];
    if Indice = 0 then
       begin
       DEV.Code:=TOBLOC.getValue('PBR_DEVISE') ;
       GetInfosDevise(DEV) ;
       end;
    TOBLOC.putValue('PBR_VALEURDEV',arrondi (TOBLOC.getValue('PBR_VALEURDEV'),DEV.decimale));
    ReAjusteMontants ( TOBLOC,DEV);
    end;
end;

procedure GetLesTaxesRG(TOBBasesRG :TOB; CatTaxe,FamTaxe:String; var RGD,RGP: double) ;
var TOBB : TOB;
begin
RGP:=0 ; RGD:=0 ;
if (TOBBasesRg=nil) then exit;
TOBB:=TOBBasesRG.findfirst(['PBR_CATEGORIETAXE','PBR_FAMILLETAXE'],[CatTaxe,FamTaxe],true);
if TOBB <> nil then
   begin
   GetLaTaxeRG (TOBB,RGD,RGP);
   end;
end;

procedure GetLaTaxeRG(TOBB :TOB; var RGD,RGP: double) ;
begin
RGP := TOBB.getValue('PBR_VALEURTAXE');
RGD := TOBB.getValue('PBR_VALEURDEV');
end;

procedure GetCumulTaxesRG(TOBBasesRG,TOBPieceRG :TOB; var RGD,RGP: double;DEV:Rdevise) ;
var TOBB : TOB;
    indice : integer;
    Ratio : double;
    XD,XP,XRD,XRP : double;
begin
RGP:=0 ; RGD:=0 ;
if (TOBBasesRg=nil) then exit;
for indice := 0 to TOBBasesRg.detail.count -1 do
    begin
    TOBB:=TOBBasesRG.detail[Indice];
    // recupere le reliquat de RG par ligne
    GetRGByIndice (TOBPieceRG, TOBB.GetValue('INDICERG'),XRP,XRD,True);
    GetRGByIndice (TOBPieceRG, TOBB.GetValue('INDICERG'),XP,XD);
    if XD <> 0 then Ratio := XRD / XD else Ratio := 0;
    RGP := RGP + (TOBB.getValue('PBR_VALEURTAXE') * ratio);
    RGD := RGD + (TOBB.getValue('PBR_VALEURDEV') * Ratio);
    end;
RGP := Arrondi(RGP,V_PGI.OkDecV );
RGD := Arrondi(RGD,DEV.decimale);
end;

procedure ReIndiceLigneRg (TOBPieceRG: TOB;NumLigne,IndiceRg: integer);
var TOBL : TOB;
begin
if TOBPieceRG = nil then exit;
TOBL:=TOBPIeceRG.findfirst(['INDICERG'],[IndiceRg],true);
while TOBL <> nil do
   begin
   TOBL.Putvalue('PRG_NUMLIGNE',Numligne);
   TOBL:=TOBPIeceRG.findnext(['INDICERG'],[IndiceRg],true);
   end;
end;

procedure ReIndiceLigneBasesRg (TOBBasesRG: TOB;NumLigne,IndiceRg: Integer);
var TOBL : TOB;
begin
if TOBBasesRG = nil then exit;
TOBL:=TOBBasesRG.findfirst(['INDICERG'],[IndiceRg],true);
while TOBL <> nil do
   begin
   TOBL.Putvalue('PBR_NUMLIGNE',Numligne);
   TOBL:=TOBBasesRG.findnext(['INDICERG'],[IndiceRg],true);
   end;
end;

procedure initLigneRg (TOBRG,TOBPIECE: TOB; Interv : string='');
begin
if TOBRG = nil then exit;
TOBRG.PutValue ('PRG_NATUREPIECEG',TOBPIECE.GetValue('GP_NATUREPIECEG'));
TOBRG.PutValue ('PRG_DATEPIECE',TOBPIECE.GetValue('GP_DATEPIECE'));
TOBRG.PutValue ('PRG_SOUCHE',TOBPIECE.GetValue('GP_SOUCHE'));
TOBRG.PutValue ('PRG_NUMERO',TOBPIECE.GetValue('GP_NUMERO'));
TOBRG.PutValue ('PRG_INDICEG',TOBPIECE.GetValue('GP_INDICEG'));
TOBRG.PutValue ('PRG_SAISIECONTRE',TOBPIECE.GetValue('GP_SAISIECONTRE'));
TOBRG.PutValue ('PRG_APPLICABLE','X');
if Interv <> '' then TOBRG.PutValue ('PRG_FOURN',Interv);
TOBRG.addchampsupValeur ('INDICERG',0);
TOBRG.AddChampSupValeur ('CAUTIONUTIL',0);
TOBRG.AddChampSupValeur ('CAUTIONUTILDEV',0);
TOBRG.AddChampSupValeur ('CAUTIONAPRES',0);
TOBRG.AddChampSupValeur ('CAUTIONAPRESDEV',0);
end;

function RestitueMontantCautionUtilise (TOBPieceRG : TOB) : boolean;
var TOBLesRG,TOBRG,TOBRGPREC : TOB;
		Q : TQuery;
    indice : integer;
    cledoc : r_cledoc;
    Interv,PiecePrec,SQL : string;
    XD,Xp : double;
begin
	result := true;
	TOBLesRG := TOB.Create ('LES RG INIT', nil,-1);
  for Indice := 0 to TOBPieceRG.detail.count -1 do
  begin
  	TOBRG := TOBPieceRG.detail[Indice];
    //
    if TOBRG.GetString('PRG_NUMCAUTION')='' then Continue;
    // Si pas de caution bancaire ce n'est pas la peine de continuer
    XD := TOBRG.GetValue('PRG_MTTTCRGDEV');
    XP := TOBRG.GetValue('PRG_MTTTCRG');
    PiecePrec := TOBRG.GetValue('PIECEPRECEDENTE');
    Interv := TOBRG.GetValue('PRG_FOURN');
		if (PiecePrec = '') or (XD =0)  then continue;
    DecodeRefPiece (PiecePrec,Cledoc);
    if cledoc.NaturePiece <> 'DBT' then Continue;  // si la provenance n'est pas celle du devis --> on ne traite pas
    SQl := 'SELECT * FROM PIECERG WHERE '+ WherePiece(CleDoc,ttdRetenuG,False)+' AND '+
    			 'PRG_NUMLIGNE='+inttostr(cledoc.NumLigne)+ ' AND '+
           'PRG_FOURN="'+Interv+'"';
    Q:=OpenSQL ( SQL,true,-1, '', True) ;
    if not Q.Eof then
    begin
      TOBRGPREC := TOB.create ('PIECERG',TOBLesRG,-1);
      TOBRGPREC.SelectDB ('',Q,true);
      TOBRGPREC.PutVAlue('PRG_CAUTIONMTU',TOBRGPREC.GetVAlue('PRG_CAUTIONMTU') - XP);
      TOBRGPREC.PutVAlue('PRG_CAUTIONMTUDEV',TOBRGPREC.GetVAlue('PRG_CAUTIONMTUDEV') - XD);
    end;
    ferme (Q);
  end;
  if TOBLesRG.detail.count > 0 then
  begin
    for Indice := 0 to TOBLesRG.detail.count -1 do
    begin
  		result := TOBLesRG.detail[Indice].UpdateDb (false);
      if not result then break;
    end;
  end;
  TOBLEsRG.free;
end;

procedure AnnuleCautionUtilise (TOBPiece,TOBPieceRG : TOB);
var TOBLESRG,TOBPRG_O,TOBRGPREC : TOB;
    XD,Xp : double;
    PieceOrigine : string;
    Q : TQuery;
    Sql : string;
    indice : integer;
    cledoc : r_cledoc;
begin
  if TOBPieceRG = nil then exit;
  XP := 0; XD := 0;

  TOBLESRG := TOB.Create ('LES RG INITIALES',nil,-1); // en provenance du devis....

  if (TOBPieceRg <> nil) then
  begin
    for indice := 0 to TOBPieceRG.detail.count -1 do
    begin
      TOBPRG_O := TOBPieceRG.detail[Indice];
      XD := TOBPRG_O.GetDouble('PRG_MTTTCRGDEV');
      XP := TOBPRG_O.GetDouble('PRG_MTTTCRG');
      PieceOrigine := TOBPRG_O.GetValue('PIECEPRECEDENTE');
      if (PieceOrigine = '') or (XD=0) then continue;
      //
      TOBRGPREC := TOB.create ('PIECERG',TOBLESRG,-1);
      DecodeRefPiece (TOBPRG_O.GetValue('PIECEPRECEDENTE'),Cledoc);
      SQl := 'SELECT * FROM PIECERG WHERE '+ WherePiece(CleDoc,ttdRetenuG,False)+' AND '+
             'PRG_NUMLIGNE='+inttostr(cledoc.NumLigne)+ ' AND '+
             'PRG_FOURN="'+ TOBPRG_O.GetString('PRG_FOURN')+'"';
      Q:=OpenSQL ( SQL,true,-1, '', True) ;
      if not Q.Eof then
      begin
        TOBRGPREC.SelectDB ('',Q,true);
      	TOBRGPREC.AddChampSupValeur ('PIECEPRECEDENTE',PieceOrigine);
        if TOBRGPREC.GetValue('PRG_CAUTIONMTDEV') <> 0 then
        begin
          TOBRGPREC.PutVAlue('PRG_CAUTIONMTU',TOBRGPREC.GetValue('PRG_CAUTIONMTU')- XP);
          TOBRGPREC.PutVAlue('PRG_CAUTIONMTUDEV',TOBRGPREC.GetValue('PRG_CAUTIONMTUDEV')-XD);
        end;
      end;
      ferme (Q);
    end;
  end;
  if TOBLESRG.detail.count > 0 then
  begin
    for Indice:= 0 to TOBLESRG.detail.count -1 do
    begin
  		TOBLEsRG.detail[Indice].UpdateDb (false);
    end;
  end;
  TOBLESRG.free;
end;

procedure ReajusteCautionDocOrigine (TOBPIece,TOBPieceRG:TOB; SensAjout : boolean;TOBPieceRG_O:TOB = nil);
var TOBL,TOBRGPREC,TOBPRG_O,TOBLESRG,TOBPRG : TOB;
    CleDoc : R_CleDoc;
    XP,XD : double;
    Q : TQuery;
    Indice : integer;
    Sql,PieceOrigine : string;
begin
  if TOBPieceRG = nil then exit;
  XP := 0; XD := 0;

  TOBLESRG := TOB.Create ('LES RG INITIALES',nil,-1); // en provenance du devis....

  if (TOBPieceRg_O <> nil) then
  begin
    for indice := 0 to TOBPieceRG_O.detail.count -1 do
    begin
      TOBPRG_O := TOBPieceRG_O.detail[Indice];
      XD := TOBPRG_O.GetDouble('PRG_MTTTCRGDEV');
      XP := TOBPRG_O.GetDouble('PRG_MTTTCRG');
      PieceOrigine := TOBPRG_O.GetValue('PIECEPRECEDENTE');
      if (PieceOrigine = '') or (XD=0) then continue;
      //
      TOBRGPREC := TOB.create ('PIECERG',TOBLESRG,-1);
      DecodeRefPiece (TOBPRG_O.GetValue('PIECEPRECEDENTE'),Cledoc);
      SQl := 'SELECT * FROM PIECERG WHERE '+ WherePiece(CleDoc,ttdRetenuG,False)+' AND '+
             'PRG_NUMLIGNE='+inttostr(cledoc.NumLigne)+ ' AND '+
             'PRG_FOURN="'+ TOBPRG_O.GetString('PRG_FOURN')+'"';
      Q:=OpenSQL ( SQL,true,-1, '', True) ;
      if not Q.Eof then
      begin
        TOBRGPREC.SelectDB ('',Q,true);
      	TOBRGPREC.AddChampSupValeur ('PIECEPRECEDENTE',PieceOrigine);
        if TOBRGPREC.GetValue('PRG_CAUTIONMTDEV') <> 0 then
        begin
          TOBRGPREC.PutVAlue('PRG_CAUTIONMTU',TOBRGPREC.GetValue('PRG_CAUTIONMTU')- XP);
          TOBRGPREC.PutVAlue('PRG_CAUTIONMTUDEV',TOBRGPREC.GetValue('PRG_CAUTIONMTUDEV')-XD);
        end;
      end;
      ferme (Q);
    end;
  end;
	for Indice := 0 to TOBPieceRG.detail.count -1 do
  begin
    TOBPRG := TOBPieceRG.detail[Indice];
    PieceOrigine := TOBPRG.GetString('PIECEPRECEDENTE');
    XD := TOBPRG.GetDouble('PRG_MTTTCRGDEV');
    XP := TOBPRG.GetDouble('PRG_MTTTCRG');
    if (Pieceorigine='') or (XD = 0) then continue;
    if PieceOrigine <> '' then
    begin
    	TOBPRG_O := TOBLESRG.findFirst(['PIECEPRECEDENTE','PRG_FOURN'],[PieceOrigine,TOBPRG.GetString('PRG_FOURN')],true);
      if TOBPRG_O = nil then
      begin
        DecodeRefPiece (PieceOrigine,Cledoc);
        SQl := 'SELECT * FROM PIECERG WHERE '+ WherePiece(CleDoc,ttdRetenuG,False)+' AND '+
               'PRG_NUMLIGNE='+inttostr(cledoc.NumLigne)+ ' AND '+
               'PRG_FOURN="'+ TOBPRG.GetString('PRG_FOURN')+'"';
        Q:=OpenSQL ( SQL,true,-1, '', True) ;
        if not Q.Eof then
        begin
      		TOBPRG_O := TOB.create ('PIECERG',TOBLESRG,-1);
          TOBPRG_O.SelectDB ('',Q,true);
          TOBPRG_O.AddChampSupValeur ('PIECEPRECEDENTE',PieceOrigine);
        end;
        ferme (Q);
      end;
      if TOBPRG_O = nil then continue;
      if TOBPRG_O.GetValue('PRG_CAUTIONMTDEV') <> 0 then
      begin
        TOBPRG_O.PutVAlue('PRG_CAUTIONMTU',TOBPRG_O.GetValue('PRG_CAUTIONMTU')+ XP);
        TOBPRG_O.PutVAlue('PRG_CAUTIONMTUDEV',TOBPRG_O.GetValue('PRG_CAUTIONMTUDEV')+XD);
      end;
    end;
  end;
  if TOBLESRG.detail.count > 0 then
  begin
    for Indice:= 0 to TOBLESRG.detail.count -1 do
    begin
  		TOBLEsRG.detail[Indice].UpdateDb (false);
    end;
  end;
  TOBLESRG.free;
end;

procedure ReinitRGUtilises (TOBPieceRG : TOB);
var TOBRg : TOB;
    Indice : integer;
begin
  if TOBPieceRG = nil then exit;
  if TOBPieceRG.FieldExists('NUMORDRE') then
  begin
  	TOBPieceRG.putvalue ('NUMORDRE',0);
  end;
  for Indice := 0 to TOBPieceRG.detail.count -1 do
  begin
    TOBRG := TOBPieceRG.detail[Indice];
		if TOBRG.FieldExists('NUMORDRE') then TOBRG.putvalue ('NUMORDRE',0);
    TOBRG.putvalue ('PRG_CAUTIONMTU',0);
    TOBRG.putvalue ('PRG_CAUTIONMTUDEV',0);
    if TOBRG.GetValue('PRG_NATUREPIECEG') = 'DBT' then    // Specif pour la V1 a parametrer plus tard
    begin
      TOBRG.PutVAlue ('PRG_NUMCAUTION','');
      TOBRG.PutVAlue ('PRG_CAUTIONMT',0);
      TOBRG.PutVAlue ('PRG_CAUTIONMTDEV',0);
    end;
  end;
end;

procedure EnregistreNumOrdreRG (TOBgenere,TOBPieceRg : TOB);
var TOBL,TOBRG : TOB;
begin
  TOBL := TOBgenere.findFirst (['GL_TYPELIGNE'],['RG'],true);
	repeat
    if TOBL = nil then break;
    if TOBL.getValue('INDICERG') <> 0 then
    begin
      TOBRG:=TOBPIeceRG.findfirst(['INDICERG'],[TOBL.getValue('INDICERG')],true);
      if TOBRG <> nil then
      begin
      	TOBRG.putValue('NUMORDRE',TOBL.GEtVAlue('GL_NUMORDRE'));
      end;
    end;
  	TOBL := TOBgenere.findNext (['GL_TYPELIGNE'],['RG'],true);
  until TOBL = nil;
end;

procedure SupprimeRg (TOBPieceRg,TOBBasesRg : TOB; IndiceRg : integer);
var TOBLOC : TOB;
begin
  if TOBPieceRg=nil then exit;
  TOBLOC:=TOBPieceRg.findfirst(['INDICERG'],[IndiceRg],true);
  While TOBLOC <> nil do
  begin
    TOBLOc.free;
    TOBLOC:=TOBPieceRg.findfirst(['INDICERG'],[IndiceRg],true);
  end;
  //
  if TOBBasesRG=nil then exit;
  TOBLOC:=TOBBasesRG.findfirst(['INDICERG'],[IndiceRg],true);
  While TOBLOC <> nil do
  begin
    TOBLOc.free;
    TOBLOC:=TOBBasesRG.findfirst(['INDICERG'],[IndiceRg],true);
  end;
end;

procedure ConstituePieceRg (TOBPiece,TOBpieceRg,TOBpieceTrait : TOB);
var TOBPT,OneTOB,TOBREFRG : TOB;
		Interv,TypePaie : string;
    Indice : integer;
begin
  if TOBPieceRG.detail.count = 0 then exit;
  TOBPieceTrait.detail.Sort('BPE_FOURNISSEUR'); // pour etre sur de traiter l'entreprise en premier
  TOBREFRG := TOBPieceRG.detail[0];
  //
  for Indice := 0 to TOBpieceTrait.detail.count -1 do
  begin
    TOBPT := TOBpieceTrait.detail[indice];
    TypePaie := TOBPT.GEtString('TYPEPAIE');
    Interv := '';
//    Interv := TOBPT.GetString('BPE_FOURNISSEUR');
//    if TOBPT.getString('TYPEPAIE')='002' then Interv := '';
    OneTOB := TOBPieceRG.findFirst(['PRG_FOURN'],[Interv],true);
    if OneTOB = nil then
    begin
      OneTOB := TOB.create('PIECERG', TOBPieceRG, -1);
      initLigneRg(OneTOB, TOBPIECE,Interv);
      OneTOB.putValue('PRG_TAUXRG',TOBREFRG.GEtValue('PRG_TAUXRG'));
      OneTOB.putValue('PRG_TYPERG',TOBREFRG.GEtValue('PRG_TYPERG'));
      OneTOB.putValue('PRG_NUMLIGNE',TOBREFRG.GEtValue('PRG_NUMLIGNE'));
      OneTOB.putValue('INDICERG',TOBREFRG.GEtValue('INDICERG'));
//      if (Interv <> '') and (TypePaie='001') then OneTOB.putValue('PRG_TAUXRG',0);
    end else
    begin
//      if (Interv <> '') and (TypePaie='001') then OneTOB.putValue('PRG_TAUXRG',0);
    end;
  end;
end;

procedure ConstitueTVARG (TOBpieceRG,TOBBASESRG,TOBBRG : TOB);

  procedure CalcBaseRGInterm (TOBBASESRG: TOB;Numligne : Integer;TOBInterm: TOB; BaseHtDev,baseHt,MtResiduHTDEV,MtREsiduHt,TvaResteDEV,TVARESTe : double);
  var II,IMAx : Integer;
      TOBI,TOBBB,TOBLL : TOB;
      ratio,Max,baseHtC,baseHtCDEv,TvaC,tvaCDEV : Double;
  begin
    Max := 0;
    TOBLL := TOB.Create ('LES LIGNES',nil,-1);
    //
    //FV1 : 09/05/2017 - FS#2533 - MORAND INDUSTRIE - Message "opération en virgule flottante incorrecte" à l'impression d'une facture
    if (MtResiduHTDEV = 0) Or (BaseHtDev = 0) then ratio := 0
    else ratio := MtResiduHTDEV / BaseHtDev;

    for II := 0 to TOBBASESRG.detail.count -1 do
    begin
      TOBBB := TOBBASESRG.detail[II];
      if TOBBB.GetInteger('PBR_NUMLIGNE')<>Numligne then Continue;
      //
      TOBI := TOB.Create ('UNE TVA RG',TOBLL,-1);
      TOBI.AddChampSupValeur('CATEGORIETAXE',TOBBB.GetString('PBR_CATEGORIETAXE'));
      TOBI.AddChampSupValeur('FAMILLETAXE',TOBBB.GetString('PBR_FAMILLETAXE'));
      TOBI.AddChampSupValeur('TAUXTAXE',TOBBB.GetString('PBR_TAUXTAXE'));
      if (TOBBB.GetValue('PBR_BASETAXE') > Max) then
      begin
        max := TOBBB.GetValue('PBR_BASETAXE');
        iMAx := TOBBB.GetIndex;
      end;
      TOBI.AddChampSupValeur('BASETAXE',ARRONDI(TOBBB.GetValue('PBR_BASETAXE')*Ratio,V_PGI.OkDecV));
      TOBI.AddChampSupValeur('VALEURTAXE',ARRONDI(TOBBB.GetValue('PBR_VALEURTAXE')*Ratio,V_PGI.OkDecV));
      TOBI.AddChampSupValeur('BASEDEV',ARRONDI(TOBBB.GetValue('PBR_BASEDEV')*Ratio,V_PGI.OkDecV));
      TOBI.AddChampSupValeur('VALEURDEV',ARRONDI(TOBBB.GetValue('PBR_VALEURDEV')*Ratio,V_PGI.OkDecV));
    end;
    baseHtC := 0; baseHtCDEV := 0; TvaC := 0; tvaCDEV := 0;
    if TOBLL.detail.count > 0 then
    begin
      for II := 0 to TOBLL.detail.count -1 do
      begin
        baseHtC := baseHtC +TOBI.getvalue('BASETAXE');
        TvaC := TvaC +TOBI.getvalue('VALEURTAXE');
        baseHtCDEV := baseHtCDEV +TOBI.getvalue('BASEDEV');
        TvaCDEV := TvaCDEV +TOBI.getvalue('VALEURDEV');
      end;
      // ---
      TOBI := TOBLL.detail[Imax];
      if MtResiduHTDEV <> baseHtCDEv then
      begin
        TOBI.PutValue('BASEDEV',ARRONDI(TOBI.GetValue('BASEDEV')-(MtResiduHTDEV-baseHTCDev),V_PGI.okdecV));
      end;
      if MtREsiduHt <> baseHtC then
      begin
        TOBI.PutValue('BASETAXE',ARRONDI(TOBI.GetValue('BASETAXE')-(MtREsiduHt-baseHtC),V_PGI.okdecV));
      end;
      if TVARESTE <> TVAC then
      begin
        TOBI.PutValue('VALEURTAXE',ARRONDI(TOBI.GetValue('VALEURTAXE')-(TVARESTE-TVAC),V_PGI.okdecV));
      end;
      if TVARESTEDEV <> tvaCDEV then
      begin
        TOBI.PutValue('VALEURDEV',ARRONDI(TOBI.GetValue('VALEURDEV')-(TVARESTEDEV-tvaCDEV),V_PGI.okdecV));
      end;

      for II := 0 to TOBLL.Detail.count -1 do
      begin
        TOBBB := TOBBRG.findfirst(['CATEGORIETAXE','FAMILLETAXE'],
                                 [TOBLL.detail[II].GetValue('CATEGORIETAXE'),
                                 TOBLL.detail[II].GetValue('CATEGORIETAXE')],true);
        if TOBBB=nil then
        begin
          TOBBB := TOB.Create ('UNE TVA RG',TOBBRG,-1);
          TOBBB.AddChampSupValeur('CATEGORIETAXE',TOBLL.detail[II].GetString('CATEGORIETAXE'));
          TOBBB.AddChampSupValeur('FAMILLETAXE',TOBLL.detail[II].GetString('FAMILLETAXE'));
          TOBBB.AddChampSupValeur('TAUXTAXE',TOBLL.detail[II].GetString('TAUXTAXE'));
          TOBBB.AddChampSupValeur('BASETAXE',0);
          TOBBB.AddChampSupValeur('VALEURTAXE',0);
          TOBBB.AddChampSupValeur('BASEDEV',0);
          TOBBB.AddChampSupValeur('VALEURDEV',0);
        end;
        TOBBB.PutValue('BASETAXE',TOBBB.GetValue('BASETAXE')+TOBLL.detail[II].GetValue('BASETAXE'));
        TOBBB.PutValue('VALEURTAXE',TOBBB.getValue('VALEURTAXE')+TOBLL.detail[II].GetValue('VALEURTAXE'));
        TOBBB.PutValue('BASEDEV',TOBBB.getValue('BASEDEV')+TOBLL.detail[II].getvalue('BASEDEV'));
        TOBBB.PutValue('VALEURDEV',TOBBB.GetValue('VALEURDEV')+TOBLL.detail[II].getValue('VALEURDEV'));
      end;
    end;
  end;

var II,Sens,NumLigne : Integer;
    TOBB : TOB;
    mtttcRG,CautionUsed,MtCaution,MtRGResiduelle,TvaReste,MtRGResiduHTDEV,BaseHt,BaseHtDev,TvaResteDEV,MtRGResiduHT : Double;
    ratio,ratioDev : double;
begin
  for II := 0 to TOBpieceRG.Detail.Count -1 do
  begin
    TOBB := TOBpieceRG.detail[II];
    if TOBB.GetString('PRG_TYPERG')<>'HT' then continue;
    ratio := 1.0;
    ratioDev := 1.0;
    BaseHtDev := TOBB.GetDouble ('PRG_MTHTRGDEV');
    BaseHt := TOBB.GetDouble ('PRG_MTHTRG');
    //
    if TOBB.GetDouble ('PRG_MTHTRGDEV') <> 0 then Ratio := TOBB.GetDouble ('PRG_MTTTCRGDEV')/ TOBB.GetDouble ('PRG_MTHTRGDEV');
    if TOBB.GetDouble ('PRG_MTTTCRG') <> 0 then RatioDev := TOBB.GetDouble ('PRG_MTTTCRGDEV')/ TOBB.GetDouble ('PRG_MTTTCRG');
    if (TOBB.GetString ('PRG_NUMCAUTION') <> '') then
    begin
      sens := 1;
      if (TOBB.GetDouble ('PRG_MTTTCRG') < 0)  then
      begin
        Sens := -1;
      end;
      //
      NumLigne := TOBB.GetInteger ('PRG_NUMLIGNE');
      mtttcRG := Abs(TOBB.GetDouble ('PRG_MTTTCRGDEV'));
      CautionUsed := Abs(TOBB.GetDouble ('CAUTIONUTILDEV'));
      MtCaution := Abs(TOBB.GetDouble ('PRG_CAUTIONMTDEV'));
      MtRGResiduelle := MtttcRG + CautionUsed - MtCaution; // TTC
      if MtRGResiduelle < 0 then MtRGResiduelle := 0;
      MtRGResiduHTDEV := Arrondi(MtRGResiduelle * sens / ratio,V_PGI.okdecV );
      MtRGResiduHT := Arrondi(MtRGResiduelle * sens / ratio,V_PGI.okdecV );
      TvaResteDEV := MtRGResiduelle - Arrondi(MtRGResiduHTDEV / ratioDev ,V_PGI.okdecV );
      TvaReste := Arrondi(TvaResteDEV/ratioDev,V_PGI.OkDecV);
      //
      CalcBaseRGInterm (TOBBASESRG,NumLigne,TOBBRG,BaseHtDev,BaseHt,MtRGResiduHTDEV,MtRGResiduHT,TvaResteDEV,TvaReste);
    end else
    begin

    end;
  end;
end;

function FindRetenue (TOBPieceRG: TOB; NumOrdre : integer) : TOB;
begin
  result :=  TOBPieceRG.findFirst(['PRG_NUMLIGNE'],[Numordre],True);
end;

function FindRGTOBLigne (TOBPiece : TOB ;NumOrdre : integer) : TOB;
var I : integer;
begin
  result := nil;
  for I := 0 to TOBPiece.detail.count -1 do
  begin
    if TOBPiece.detail[I].GetString('GL_TYPELIGNE') <> 'RG' then continue;
    if TOBPiece.detail[I].GetInteger('GL_NUMORDRE')=NumOrdre then
    begin
      result := TOBPiece.detail[I];
    end;
  end;
end;

end.
