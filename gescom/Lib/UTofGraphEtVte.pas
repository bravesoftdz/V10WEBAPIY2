unit UTofGraphEtVte;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, Hqry, HSysMenu, hmsgbox, Menus, ComCtrls, TeeProcs,Spin,
  UtilSynVte,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      Fiche, Fe_Main,
{$ENDIF}
  GRS1, TeEngine, Chart, Grids, Hctrls, ExtCtrls, StdCtrls, HTB97, UIUtil, UTOF,
  UTOB, HStatus, HPanel, M3FP, Hent1, EntGC, GraphUtil, Series, UtilArticle, Ent1;

Type
     TOF_GraphEtVte = Class (TOF)
        procedure OnArgument (Argument : string); override;
        procedure OnUpdate; override;
        procedure ChargeGraph;
        procedure OnClose  ; override ;
     private
        TobSyn : TOB;
        datdeb, datfin : TDateTime ;
     end;

implementation

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_GraphEtVte.OnArgument (Argument : string);
var F : TForm;

begin
inherited;
F := TForm (Ecran);
// modif le 13/08/02 prise en compte de l'établissement par défaut et non du dépôt par défaut
THValComboBox(F.FindComponent ('ETABLISSEMENT')).Value := VH^.EtablisDefaut; //VH_GC.GCDepotDefaut;
TToolbarButton97(F.FindComponent('bAffGraph')).Down:=True;
THGrid(F.FindComponent('FListe')).Visible:=False;
end;

procedure TOF_GraphEtVte.OnUpdate;
Var F : TFGRS1;
    stWhere, ston : string;
    QLigne : TQuery;
    TobTemp : TOB;
    Edit : THEdit ;
    Edit2 : THValcombobox ;
    Edit4 : TSpinEdit ;
    nbmois : integer ;
    ctrl, ctrlName,signe,rupt1,rupt2 : string ;
begin
inherited;
F := TFGRS1(Ecran);
ExecuteSQL('DELETE FROM GCTMPSYNVTE WHERE GZS_UTILISATEUR = "'+V_PGI.USer+'" AND GZS_TRAIT = "ETV"');

ctrl:='GQ_DEPOT';
ctrlName:='ETABLISSEMENT'; signe:='=' ;
Edit2:=THValComboBox(TFGRS1(F).FindComponent(ctrlName)) ;
if (Edit2<>nil) and (Edit2.Value<>'') and (Edit2.Value<>'<<Tous>>') then
  begin
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
  end ;

ctrl:='GA_CODEARTICLE';
ctrlName:='CODEARTICLE'; signe:='=' ;
Edit:=THEdit(TFGRS1(F).FindComponent(ctrlName)) ;
if (Edit<>nil) and (Edit.Text<>'') then
  begin
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit.Text+'"' ;
  end ;

ctrl:='GA_FOURNPRINC';
ctrlName:='FOURNPRINC'; signe:='=' ;
Edit:=THEdit(TFGRS1(F).FindComponent(ctrlName)) ;
if (Edit<>nil) and (Edit.Text<>'') then
  begin
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit.Text+'"' ;
  end ;

ctrl:='GA_COLLECTION';
ctrlName:='COLLECTION'; signe:='=' ;
Edit2:=THValComboBox(TFGRS1(F).FindComponent(ctrlName)) ;
if (Edit2<>nil) and (Edit2.Value<>'') and (Edit2.Value<>'<<Toutes>>') then
  begin
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
  end ;

ctrl:='GA_FAMILLENIV1';
ctrlName:='FAMILLENIV1'; signe:='=' ;
Edit2:=THValComboBox(TFGRS1(F).FindComponent(ctrlName)) ;
if (Edit2<>nil) and (Edit2.Value<>'') and (Edit2.Value<>'<<Toutes>>') then
  begin
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
  end ;

ctrl:='GA_FAMILLENIV2';
ctrlName:='FAMILLENIV2'; signe:='=' ;
Edit2:=THValComboBox(TFGRS1(F).FindComponent(ctrlName)) ;
if (Edit2<>nil) and (Edit2.Value<>'') and (Edit2.Value<>'<<Toutes>>') then
  begin
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
  end ;

ctrl:='GA_FAMILLENIV3';
ctrlName:='FAMILLENIV3'; signe:='=' ;
Edit2:=THValComboBox(TFGRS1(F).FindComponent(ctrlName)) ;
if (Edit2<>nil) and (Edit2.Value<>'') and (Edit2.Value<>'<<Toutes>>') then
  begin
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
  end ;

ctrl:='GL_DATEPIECE';
Edit:=THEdit(TFGRS1(F).FindComponent('DATJOUR')) ;
datfin := StrToDate(Edit.Text);
datfin := FinDeMois(datfin);
edit4 := TSpinEdit(TFGRS1(F).FindComponent('NBMOIS'));
nbmois := edit4.value;
datdeb:=PlusMois(datfin,-nbmois) ;
datdeb := DebutDeMois(datdeb);
if TcheckBox(TFGRS1(F).FindComponent('ENCOURS')).State <> cbChecked then
       datfin:=PlusMois(datfin,-1) ;
ston:=' AND '+ctrl+' >= "'+USDateTime(datdeb)+'"' ;
ston:=ston+' AND '+ctrl+' <= "'+USDateTime(datfin)+'"' ;

QLigne := OpenSQL('SELECT GA_CODEARTICLE,GA_LIBELLE, GA_COLLECTION,GA_FOURNPRINC,'+
'GA_FAMILLENIV1,GA_FAMILLENIV2,GA_FAMILLENIV3, GQ_DEPOT, '+
'GL_DATEPIECE, GL_QTEFACT,GQ_PHYSIQUE as STOCK,(GQ_RESERVEFOU - GQ_LIVREFOU) as ATTENDU, '+
'CC1.CC_LIBELLE AS LIBFAMNIV1, CC2.CC_LIBELLE AS LIBFAMNIV2, CC3.CC_LIBELLE AS LIBFAMNIV3, '+
'CC4.CC_LIBELLE AS COLLEC, T_LIBELLE, ET_LIBELLE FROM DISPOART_MODE '+
'LEFT JOIN LIGNE ON GQ_DEPOT = GL_ETABLISSEMENT AND GA_CODEARTICLE = GL_CODEARTICLE '+
' AND GL_NATUREPIECEG="FFO" '+ston+
' LEFT OUTER JOIN CHOIXCOD CC1 ON GA_FAMILLENIV1=CC1.CC_CODE AND CC1.CC_TYPE="FN1"'+
' LEFT OUTER JOIN CHOIXCOD CC2 ON GA_FAMILLENIV2=CC2.CC_CODE AND CC2.CC_TYPE="FN2"'+
' LEFT OUTER JOIN CHOIXCOD CC3 ON GA_FAMILLENIV3=CC3.CC_CODE AND CC3.CC_TYPE="FN3"'+
' LEFT OUTER JOIN CHOIXCOD CC4 ON GA_COLLECTION=CC4.CC_CODE AND CC4.CC_TYPE="GCO" '+
' LEFT OUTER JOIN TIERS ON GA_FOURNPRINC=T_TIERS '+
' LEFT OUTER JOIN ETABLISS ON GQ_DEPOT=ET_ETABLISSEMENT '+
'WHERE GA_TYPEARTICLE = "MAR" AND GQ_CLOTURE = "-" AND '+stWhere+
'ORDER BY GQ_DEPOT,GA_CODEARTICLE,GL_DATEPIECE ', True);

if QLigne.Eof then
   begin
   Ferme(QLigne);
   exit;
   end;

TobTemp:=TOB.Create('',nil,-1);
TobTemp.LoadDetailDB('','','',QLigne,false);
Ferme(QLigne);

TobSyn := TOB.Create('TMPSYNVTE',nil,-1);
Edit2:=THValComboBox(TFGRS1(F).FindComponent('REGROUP')) ;
rupt2 := Edit2.Value;
Edit2:=THValComboBox(TFGRS1(F).FindComponent('ORDER')) ;
rupt1 := Edit2.Value;
TobSyn.Detail.Sort('RUPT1;RUPT2;ANNEE;MOIS') ;
RempliTOBSynthese(TobTemp, TobSyn, 'AA', True, False,'GRV',rupt2,rupt1, datdeb, datfin);
TobTemp.Free;
ChargeGraph;
TobSyn.Free;
end;

procedure TOF_GraphEtVte.OnClose  ;
begin
//TobSyn.Free;
end;

procedure TOF_GraphEtVte.ChargeGraph;
var F : TFGRS1;
    sttitre, stColonnes, stChampTitre, stCriteres : string;
    tstTitres : Tstrings;
    stTitresCol1, stTitresCol2, stColGraph1, stColGraph2 : string;
begin
stCriteres := '';
stColGraph1 := '';
stColGraph2 := '';
stTitresCol1 := '';
stTitresCol2 := '';
sttitre := '';

tstTitres := TStringList.Create;
F := TFGRS1(Ecran);

//stColonnes := 'GZS_MOIS;GZS_QTESTOCK;GZS_QTEVTE';
//stTitre := 'Mois;Stock;Quantité vendue';            // légende des colonnes
//stColGraph1 := 'GZS_QTESTOCK;GZS_QTEVTE';
//stColonnes := 'GZS_MOIS;GZS_QTESTOCK;GZS_QTEVTE';
//stTitre := 'Mois;Stock;Quantité vendue';            // légende des colonnes

tstTitres.Add ('Etat des ventes mensuelles');
stColonnes := 'GZS_MOIS;GZS_QTEVTE';
stTitre := 'Mois;Quantité vendue';            // légende des colonnes
stColGraph1 := 'GZS_QTEVTE';
stChampTitre := stTitre;

LanceGraph (F, TobSyn, '', stColonnes,  '', stTitre ,
            stColGraph1, stColGraph2, tstTitres, nil, TBarSeries, stChampTitre, False);
F.FChart1.SeriesList.series[0].Marks.Visible := False;
//F.FChart1.SeriesList.series[1].Marks.Visible := False;
//F.FChart1.SeriesList.series[0].SeriesColor := clLime;
//F.FChart1.SeriesList.series[1].SeriesColor := clRed;
F.FChart1.LeftAxis.Title.caption := 'Quantité';
F.FChart1.BottomAxis.Title.caption := 'Mois';
tstTitres.Free;
end;


Procedure TOF_GraphEtVte_ChangeAffichage (parms:array of variant; nb: integer );
var F : TFGRS1;
     TOTOF : TOF ;
begin
F:=TFGRS1(Longint(Parms[0])) ;
if (F.Name <> 'GCGraphEtVte') then exit;
if (F is TFGRS1) then TOTOF:=TFGRS1(F).Latof else exit;
if (TOTOF is TOF_GraphEtVte) then
  TOF_GraphEtVte(TOTOF).ChargeGraph else exit;
end;

procedure InitTOFGraphEtVte ();
begin
RegisterAglProc('GraphEtVte_ChangeAffichage', True , 1, TOF_GraphEtVte_ChangeAffichage);
end;

Initialization
registerclasses([TOF_GraphEtVte]);
InitTOFGraphEtVte();
end.
