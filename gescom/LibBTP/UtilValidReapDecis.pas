unit UtilValidReapDecis;

interface
uses {$IFDEF VER150} variants,{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist,Hctrls,
{$IFDEF EAGLCLIENT}
      maineagl,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     UTofListeInv,
     FactCommBtp,
     EntGc,
{$ENDIF}
     Mask, HPanel,UTOB,HEnt1, FactTOB, FactPiece,FactComm,
     ParamSoc, FactContreM,uEntCommun,UtilTOBPiece ;

procedure ConstitueLaTOBFinale_UtilVRD (TOBLigneSel: TOB; ChampsTri,RuptureDoc,ruptureArt,Mode : string; NbjourSecu : integer);
procedure GenereLATobSortie_UtilVRD (TOBIntermediaire,TOBLigneSel : TOB; Mode : string; NbjourSecu : integer);
procedure TrielaTOB_UtilVRD(TOBLigneSel: TOB ; ParQuoi : string);
function ConstitueChampsTriRupt_UtilVRD(var Sort: string; var RuptDoc : string; Var RuptArt : string; EclatCdeFou : boolean): string;  //a la place de tri j'ai faillit ecrire Tage

implementation


{***********A.G.L.Privé.*****************************************
Auteur  ...... : SANTUCCI
Créé le ...... : 20/05/2005
Modifié le ... :   /  /    
Description .. : procedure permettant de constituer une Tob en vu de 
Suite ........ : generer des cdes fournisseurs avec gestion des 
Suite ........ : regroupements.
Mots clefs ... : 
*****************************************************************}
procedure ConstitueLaTOBFinale_UtilVRD (TOBLigneSel: TOB; ChampsTri,RuptureDoc,ruptureArt,Mode : string; NbjourSecu : integer);

	function GetValeurElt(TOBL : TOB;lesChamps : string) : string;
  var UnChamps,UneValeur : string;
  begin
  	result := '';
    repeat
    	unChamps := readTokenSt (LesChamps);
      if TOBL.fieldExists (unChamps) then
      begin
      	result := result + VarasType(TOBL.getValue (UnChamps),varString)+';';
      end;
    until unchamps = '';
  end;

var TOBIntermediaire : TOB;
		TOBL,TOBLD,TOBDest,TOBDArt,TOBLART,TOBLReg,TOBDoc : TOB;
    Indice,Detail,NumOrdre,IArticle,Document: integer;
    TraitementDoc,TraitementArt,ComparaisonSort : string;
    ThisDoc,ThisArt,CurrentSort : string;
begin
	TOBIntermediaire := TOB.Create ('THE TOB INTERMEDIAIRE',nil,-1);
  TraitementDoc := '';
  TraitementArt := '';
  ComparaisonSort := '';
  Document := 1;
  TRY
    //
    NumOrdre := 1;
    for Indice := 0 to TOBLigneSel.detail.count -1 do
    begin
    	TOBL := TOBLigneSel.detail[Indice];
    	// Traitement de la rupture
      ThisDoc := GetValeurElt(TOBL,RuptureDoc);
      ThisArt := GetValeurElt(TOBL,RuptureArt);
      CurrentSort := GetValeurElt(TOBL,ChampsTri);
      if (TraitementDoc <> ThisDoc) then
      begin
      	// Emulation de la Creation du document
        TOBDest := TOB.Create ('UN DOCUMENT SORTIE',TOBIntermediaire,-1);
        TOBDest.AddChampSupValeur ('DOCUMENT',ThisDoc,false);
        TraitementDoc := ThisDoc;
        TraitementArt := '';
      end;
      if TraitementArt <> ThisArt then
      begin
        TOBDART := TOB.Create ('UN ARTICLE',TOBDest,-1);
        TOBDART.AddChampSupValeur ('ARTICLE',ThisArt,false);
        TraitementArt := ThisArt;
				TOBDART.AddChampSupValeur ('NUMRATACHEMENT',NumOrdre);
        inc(NumOrdre);
      end;
      TOBLART := TOB.Create ('LIGNE',TOBDART,-1);
      TOBLArt.Dupliquer (TOBL,true,true);
    end;
    //
    GenereLATobSortie_UtilVRD (TOBIntermediaire,TOBLigneSel,Mode,nbJourSecu);
  FINALLY
  	TOBInterMediaire.free;
  END;
end;

{****************************************************************
Auteur  ...... : SANTUCCI
Créé le ...... : 20/05/2005
Modifié le ... :   /  /
Description .. : procedure permettant de générer une tob de ligne préparées
Suite ........ : pour générer des cdes fournisseurs
Mots clefs ... : DECISIONACH;REAPPRO;VALIDATION
*****************************************************************}
procedure GenereLATobSortie_UtilVRD (TOBIntermediaire,TOBLigneSel : TOB; Mode : string; NbjourSecu : integer);

  function IsMultiAffaire (LATOB : TOB) : boolean;
  var Indice : integer;
  		LaRef : string;
      ThisTOB : TOB;
  begin
  	result := false;
    for Indice := 0 to LaTOB.detail.count -1 do
    begin
    	ThisTOB := LaTOB.detail[Indice];
      if Indice = 0 then
      begin
      	LaRef := ThisTOB.GetValue('GL_AFFAIRE');
      end;
      if LaRef <> ThisTOB.GetValue('GL_AFFAIRE') then
      begin
      	result := true;
        break;
      end;
    end;
  end;

var NbrDoc,NbrArt,DetARt : integer;
		TOBDoc,TOBDDOC,TOBLD,TOBART,TOBLReg,TOBLART : TOB;
    Cledoc : R_cledoc;
    DateL : TdateTime;
begin
  TOBLigneSel.ClearDetail;
  // On réécrit dans la toblignesel avec les regroupements qui vont bien
  for NbrDoc := 0 to TOBIntermediaire.detail.count -1 do
  begin
    TOBDOC := TOBIntermediaire.detail[NbrDoc]; // Un document a traiter
    for NbrArt := 0 To TOBDOC.Detail.count -1 do
    begin
      TOBART := TOBDOC.Detail[NbrArt]; // Ligne de regroupement tOB intermediaire
      if TOBART.Detail.Count > 1 then
      begin
        TOBLReg := TOB.Create ('LIGNE',TOBLigneSel,-1); // Ca c'est la ligne de regroupement sur document destination
        for DetARt := 0 to TOBART.detail.count -1 do
        begin
          TOBLD := TOBART.detail[DetARt];
          if Detart = 0 then
          begin
            TOBLReg.Dupliquer (TOBLD,true,true);
            AddLesSupLigne (TOBLReg,false);
            TOBLReg.AddChampSupValeur ('NUMRATACHEMENT',TOBART.GetValue('NUMRATACHEMENT'));
            TOBLReg.AddChampSupValeur ('DOCUMENT',TOBDOC.GetValue('DOCUMENT'));
            if ((Mode = 'REAP') OR (Mode = 'VALIDDECIS')) and (IsMultiAffaire(TOBLD)) then
            begin
              TOBLReg.putValue('GL_AFFAIRE','');
              TOBLReg.putValue('GL_AFFAIRE1','');
              TOBLReg.putValue('GL_AFFAIRE2','');
              TOBLReg.putValue('GL_AFFAIRE3','');
              TOBLReg.putValue('GL_AVENANT','');
            end;
            TOBLReg.putValue('GL_PIECEPRECEDENTE','');
            DecodeRefPiece (TOBLReg.GetValue('GL_PIECEORIGINE'),cledoc);
            cledoc.NumLigne := 0;
            cledoc.NumOrdre := 0;
            TOBLReg.putValue('GL_PIECEORIGINE',CleDocToString (cledoc));
//            TOBLReg.putValue('GL_PIECEORIGINE','');
            if TOBLReg.FieldExists ('PIECEORIGINE') then  TOBLReg.DelChampSup ('PIECEORIGINE',false);
            TOBLReg.putValue('GL_TYPELIGNE','CEN');
            TOBLReg.putValue('GL_LIGNELIEE',0);
            TOBLReg.putValue('GL_QTEFACT',0);
            TOBLReg.putValue('GL_TENUESTOCK','-');
          	TOBLReg.PutValue ('NUMRATACHEMENT',TOBART.GetValue('NUMRATACHEMENT'));
            TOBLReg.putValue('GL_LIGNELIEE',TOBLReg.GetValue('NUMRATACHEMENT'));
            //
            TOBLReg.putValue('GL_COEFFG',TOBLD.GetValue('GL_COEFFG'));
            TOBLReg.putValue('GL_COEFMARG',TOBLD.GetValue('GL_COEFMARG'));
            TOBLReg.putValue('GL_PUHTDEV',TOBLD.GetValue('GL_PUHTDEV'));
            //
      			DateL := PlusDate (TOBLReg.GetDateTime('GL_DATELIVRAISON'),(NbJourSecu*-1),'J');
          	if DateL < Date then DateL := Date;
            TOBLReg.PutValue('GL_DATELIVRAISON', DateL);
  					//
          end;

          TOBLArt := TOB.Create ('LIGNE',TOBLIgneSel,-1); // Et Là la ligne detail
          TOBLArt.dupliquer (TOBLD,true,true);
          TOBLART.AddChampSupValeur ('RATACHE',TOBArt.GEtValue('NUMRATACHEMENT'));
          TOBLART.AddChampSupValeur ('DOCUMENT',TOBDOC.GetValue('DOCUMENT'));
          TOBLART.putValue('GL_LIGNELIEE',TOBLART.GetValue('RATACHE'));
      		DateL := PlusDate (TOBLART.GetDateTime('GL_DATELIVRAISON'),(NbJourSecu*-1),'J');
          if DateL < Date then DateL := Date;
          TOBLART.PutValue('GL_DATELIVRAISON', DateL);
          // Cumul
          TOBLReg.putValue ('GL_QTEFACT',TOBLReg.GetValue ('GL_QTEFACT')+TOBLART.getValue('GL_QTEFACT'));
          TOBLReg.putValue ('GL_QTESTOCK',TOBLReg.GetValue ('GL_QTEFACT'));
          TOBLReg.putValue ('GL_QTERESTE',TOBLReg.GetValue ('GL_QTEFACT'));
          TOBLReg.putValue ('GL_MTRESTE',TOBLReg.GetValue ('GL_MONTANTHTDEV'));
        end;
      end else
      begin
      	TOBLD := TOBART.detail[0];
        TOBLArt := TOB.Create ('LIGNE',TOBLIgneSel,-1); // Et La la ligne detail
        TOBLArt.dupliquer (TOBLD,true,true);
        TOBLART.AddChampSupValeur ('DOCUMENT',TOBDOC.GetValue('DOCUMENT'));
      	DateL := PlusDate (TOBLArt.GetDateTime('GL_DATELIVRAISON'),(NbJourSecu*-1),'J');
        if DateL < Date then DateL := Date;
        TOBLArt.PutValue('GL_DATELIVRAISON', DateL);
      end;
    end;
  end;
end;


function ConstitueChampsTriRupt_UtilVRD(var Sort: string; var RuptDoc : string; Var RuptArt : string; EclatCdeFou : boolean): string;  //a la place de tri j'ai faillit ecrire Tage
var fRegroupLibellediff : boolean;
begin
  fRegroupLibellediff := GetParamSocSecur('SO_REGROUPLIBELLEDIFF',false);
	IF EclatCdeFou then
  begin
    if fRegroupLibellediff then
    begin
      Sort := 'GL_TIERS;GL_ETABLISSEMENT;GL_DEPOT;GL_AFFAIRE;GL_IDENTIFIANTWOL;GL_DATELIVRAISON;GL_ARTICLE';
    end else
    begin
      Sort := 'GL_TIERS;GL_ETABLISSEMENT;GL_DEPOT;GL_AFFAIRE;GL_IDENTIFIANTWOL;GL_DATELIVRAISON;GL_ARTICLE;GL_LIBELLE';
    end;
    RuptDoc := 'GL_TIERS;GL_ETABLISSEMENT;GL_DEPOT;GL_AFFAIRE;GL_IDENTIFIANTWOL;';
    RuptArt := Sort;
  end else
  begin
    Sort := 'GL_TIERS;GL_ETABLISSEMENT;GL_DEPOT;GL_DATELIVRAISON;GL_ARTICLE;GL_IDENTIFIANTWOL;GL_AFFAIRE';
    RuptDoc := 'GL_TIERS;GL_ETABLISSEMENT;';
    if fRegroupLibellediff then
    begin
      RuptArt := 'GL_TIERS;GL_ETABLISSEMENT;GL_DEPOT;GL_DATELIVRAISON;GL_ARTICLE;GL_IDENTIFIANTWOL;';
    end else
    begin
      RuptArt := 'GL_TIERS;GL_ETABLISSEMENT;GL_DEPOT;GL_DATELIVRAISON;GL_ARTICLE;GL_LIBELLE;GL_IDENTIFIANTWOL;';
    end;
  end;
end;

procedure TrielaTOB_UtilVRD(TOBLigneSel: TOB ; ParQuoi : string);
begin
	TOBLigneSel.Detail.Sort(ParQuoi);
end;


end.
