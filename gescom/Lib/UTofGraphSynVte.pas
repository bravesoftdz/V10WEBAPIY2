unit UTofGraphSynVte;

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
     TOF_GraphSynVte = Class (TOF)
        procedure OnArgument (Argument : string); override;
        procedure OnUpdate; override;
        procedure regroupe_ligne(TobSyn : TOB; var TobTemp : TOB);
        procedure regroupe_ligne_jour(TobSyn : TOB ; var TobTemp : TOB);
        procedure Cumul_ligne(var TobTemp : TOB);
     private
        stcriteres : string;
        Critere : string;
        Prixttc, Prixach, Cumprog, compar : boolean;
        annee_dem, mois_dem : word;
     end;

implementation

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_GraphSynVte.OnArgument (Argument : string);
var F : TForm;
    stArgument : string;
    Spin : TSpinEdit;
begin
inherited;
F := TForm (Ecran);
//modif le 13/08/02 car ici c'est la notion d'établissement qu'il faut et non celui du dépôt (pour que ça marche en multi-dépôt) 
THValComboBox(F.FindComponent ('ETABLISSEMENT')).Value := VH^.EtablisDefaut; //VH_GC.GCDepotDefaut;
TToolbarButton97(F.FindComponent('bAffGraph')).Down:=True;
THGrid(F.FindComponent('FListe')).Visible:=False;
stArgument := Argument;
Critere:=uppercase(Trim(ReadTokenSt(stArgument))) ;
if critere = 'AD'  then  Ecran.Caption:='Synthèse Annuelle';
if critere = 'MD' then Ecran.Caption:='Synthèse Mensuelle';
updatecaption(Ecran);
Spin := TSpinEdit(F.FindComponent('EDTANNEE'));
if Spin <> nil then Spin.value := strtoInt (FormatDateTime('yyyy', Date));
Spin := TSpinEdit(F.FindComponent('EDTMOIS'));
if Spin <> nil then Spin.value := strtoInt (FormatDateTime('mm', Date));
end;

procedure TOF_GraphSynVte.OnUpdate;
Var F : TFGRS1;
    titre, stColonnes, stChampTitre : string;
    tstTitres : Tstrings;
    stTitresCol1, stTitresCol2, stColGraph1, stColGraph2 : string;
    stWhere : string;
    QEtiq : TQuery;
    TobTemp, TobSyn : TOB;
    Edit : THEdit ;
    Edit2 : THValcombobox ;
    Edit3 : TSpinEdit ;
    iBorne, nbjour : integer ;
    ctrl, ctrlName,signe: string ;
    LaDate : TDateTime ;
    Jour : Word;
begin
inherited;
F := TFGRS1(Ecran);
stCriteres := '';
stColGraph1 := '';
stColGraph2 := '';
stTitresCol1 := '';
stTitresCol2 := '';
titre := '';

tstTitres := TStringList.Create;

Prixttc := False;
Prixach := False;
Cumprog := False;
compar := False;
if TcheckBox(TFGRS1(F).FindComponent('EDTTTC')).State = cbChecked then
   Prixttc := True;
if TcheckBox(TFGRS1(F).FindComponent('EDTCUM')).State = cbChecked then
   Cumprog := True;
if TcheckBox(TFGRS1(F).FindComponent('COMPAR')).State = cbChecked then
   compar := True;
ExecuteSQL('DELETE FROM GCTMPSYNVTE WHERE GZS_UTILISATEUR = "'+V_PGI.USer+'" AND GZS_TRAIT = "SYN"');
edit3 := TSpinEdit(TFGRS1(F).FindComponent('EDTANNEE'));
annee_dem := edit3.Value;
ctrl:='GL_ETABLISSEMENT';
ctrlName:='ETABLISSEMENT'; signe:='=' ;
Edit2:=THValComboBox(TFGRS1(F).FindComponent(ctrlName)) ;
if (Edit2<>nil) and (Edit2.Value<>'') and (Edit2.Value<>'<<Tous>>') then
  begin
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
  end ;

ctrl:='GL_DATEPIECE';
if Ecran.Name = 'GCGRAPHSYNM' then
  begin
  Edit:=THEdit(TFGRS1(F).FindComponent('DATEPIECE')) ;
  Ladate := StrToDate(Edit.Text);
  DecodeDate(LaDate, annee_dem, mois_dem, Jour);
  Jour := 1;
  nbjour := 30;
  if (mois_dem = 1) or (mois_dem = 3) or (mois_dem = 5)
  or (mois_dem = 7) or (mois_dem = 8) or (mois_dem = 10)
  or (mois_dem = 12) then nbjour := 31;
  if (mois_dem = 2) then
    begin
    if (isLeapYear(annee_dem) = True) then nbjour := 29
    else nbjour := 28;
    end;
  LaDate := Encodedate (Annee_dem, Mois_dem, Jour);
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+'(('+ctrl+' >= "'+USDateTime(Ladate)+'"' ;
  Jour := nbjour;
  LaDate := Encodedate (Annee_dem, Mois_dem, Jour);
  if stWhere<>'' then stWhere:=stWhere+' AND ' ;
  stWhere:=stWhere+ctrl+' <= "'+USDateTime(Ladate)+'" )' ;
  if (compar = True) then
    begin
    Jour := 1;
    if (mois_dem = 2) then
      begin
      if (isLeapYear(annee_dem) = True) then nbjour := 29
      else nbjour := 28;
      end;
    LaDate := Encodedate (Annee_dem -1, Mois_dem, Jour);
    if stWhere<>'' then stWhere:=stWhere+' OR (' ;
    stWhere:=stWhere+ctrl+' >= "'+USDateTime(Ladate)+'"' ;
    Jour := nbjour;
    LaDate := Encodedate (Annee_dem -1, Mois_dem, Jour);
    if stWhere<>'' then stWhere:=stWhere+' AND ' ;
    stWhere:=stWhere+ctrl+' <= "'+USDateTime(Ladate)+'" )' ;
    end;
  stWhere:=stWhere+')' ;
  end;

if Ecran.Name = 'GCGRAPHSYNA' then
  begin
  for iBorne:=1 to 2 do
    begin
    if iBorne=1
    then begin ctrlName:='DATEPIECE_DEB'; signe:='>=' ; end
    else begin ctrlName:='DATEPIECE_FIN' ; signe:='<=' ; end ;
    Edit:=THEdit(TFGRS1(F).FindComponent(ctrlName)) ;
    if (Edit<>nil) and (Edit.Text<>'01/01/1900') and (Edit.Text<>'31/12/2099') then
      begin
      if stWhere<>'' then stWhere:=stWhere+' AND ' ;
      stWhere:=stWhere+ctrl+signe+'"'+USDateTime(StrToDate(Edit.Text))+'"' ;
      end ;
    end ;
  end;

QEtiq := OpenSQL('SELECT GL_ETABLISSEMENT,GL_NATUREPIECEG,GL_DATEPIECE,'+
'GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE,GL_DEVISE,GL_QTEFACT,GL_TOTALHT,'+
'GL_TOTALTTC,GL_PMAP,(GL_PUTTC*GL_QTEFACT)/GL_PRIXPOURQTE as BRUTTC, '+
'(GL_PUHT*GL_QTEFACT)/GL_PRIXPOURQTE as BRUTHT FROM LIGNE'+
' WHERE GL_NATUREPIECEG="FFO" AND '+stWhere+' ORDER BY GL_ETABLISSEMENT,'+
' GL_DATEPIECE,GL_SOUCHE,GL_NUMERO,GL_INDICEG,GL_NUMLIGNE', True);

if QEtiq.Eof then
   begin
   HShowMessage('0;Pas de vente sur cette période;;E;O;O;O;','','') ;
   Ferme(QEtiq);
   exit;
   end;

TobTemp:=TOB.Create('',nil,-1);
TobTemp.LoadDetailDB('','','',QEtiq,false);
Ferme(QEtiq);
TobSyn := TOB.Create('GCTMPSYNVTE',nil,-1);
if (critere = 'MD') then
  RempliTousLesJours(TobTemp, TobSyn);
RempliTOBSynthese(TobTemp, TobSyn, critere, Prixttc, Prixach,'SYN','','',0,0);
TobTemp.Free;
TobTemp:=TOB.Create('EDTSYN',nil,-1);
if (critere = 'AD') then
  regroupe_ligne(TobSyn, TobTemp);
if (critere = 'MD') then
  regroupe_ligne_jour(TobSyn, TobTemp);
// si cumul prograssif demandé, modification des totaux de la TOB
if (cumprog = True) then
  Cumul_ligne(TobTemp);
if (compar = True) then
  begin
  if (critere = 'AD') then
    begin
    stColonnes := 'LIBMOIS;CATOT_ANT;CATOT';
    Titre := 'Mois;C.A. '+inttostr(annee_dem -1)+';C.A. '+inttostr(annee_dem)+'';            // légende des colonnes
    tstTitres.Add ('Synthèse des ventes Annuelle');
    end;
  if (critere = 'MD') then
    begin
    stColonnes := 'JOUR;CATOT_ANT;CATOT';
    Titre := 'Jour;C.A. '+inttostr(annee_dem -1)+';C.A. '+inttostr(annee_dem)+'';            // légende des colonnes
    tstTitres.Add ('Synthèse des ventes Mensuelle');
    end;
  stColGraph1 := 'CATOT_ANT;CATOT';
  stChampTitre := '';
end else
begin
  if (critere = 'AD') then
    begin
    stColonnes := 'LIBMOIS;CATOT';
    Titre := 'Mois;C.A.';            // légende des colonnes
    tstTitres.Add ('Synthèse des ventes Annuelle');
    end;
  if (critere = 'MD') then
    begin
    stColonnes := 'JOUR;CATOT';
    Titre := 'Jour;C.A.';            // légende des colonnes
    tstTitres.Add ('Synthèse des ventes Mensuelle');
    end;
  stColGraph1 := 'CATOT';
  stChampTitre := '';
end;
//stChampTitre := 'GZS_MOIS;GZS_CATOT;GZS_MTREELHT';

LanceGraph (F, TobTemp, '', stColonnes,  '', Titre ,
            stColGraph1, stColGraph2, tstTitres, nil, TBarSeries, stChampTitre, False);
F.FChart1.SeriesList.series[0].Marks.Visible := False;
if (compar = True) then
  begin
  F.FChart1.SeriesList.series[1].Marks.Visible := False;
  F.FChart1.SeriesList.series[0].SeriesColor := clGreen;
  F.FChart1.SeriesList.series[1].SeriesColor := clRed;
  end;
F.FChart1.LeftAxis.Title.caption := V_PGI.DevisePivot;
if (critere = 'AD') then
  F.FChart1.BottomAxis.Title.caption := 'Mois';
if (critere = 'MD') then
  F.FChart1.BottomAxis.Title.caption := 'Jours';

//F.FChart1.OnClickSeries:=ZoomGraphe;
TobSyn.Free;
TobTemp.Free;

tstTitres.Free;
end;

procedure TOF_GraphSynVte.regroupe_ligne(TobSyn : TOB ; var TobTemp : TOB);
var TobPal,TobT : TOB;
    i_ind : integer;
    cum, cum_ant : double;
begin
// génération des 12 mois de l'année dans la TOB
for i_ind:=1 to 12 do
  begin
  TobPal := TOB.Create('edtsyn',TobTemp,-1);
  TobPal.AddChampSup('MOIS',False);
  TobPal.AddChampSup('LIBMOIS',False);
  TobPal.AddChampSup('CATOT',False);
  TobPal.AddChampSup('CATOT_ANT',False);
  TobPal.InitValeurs;
  TobPal.PutValue('MOIS',i_ind);
  TobPal.PutValue('LIBMOIS',ShortMonthNames[i_ind]);
  TobPal.PutValue('CATOT',0);
  TobPal.PutValue('CATOT_ANT',0);
  end;

// chargement des cumul pour chaque mois
InitMove(TobSyn.Detail.count,'');
for i_ind:=0 to TobSyn.Detail.count-1 do
  begin
  TobT := TobSyn.Detail[i_ind];
  TobPal := TobTemp.FindFirst(['MOIS'],[IntToStr(TobT.GetValue('GZS_MOIS'))], False);
  MoveCur(False);
  if TobPal <> nil then
  begin
    if TobT.GetValue('GZS_ANNEE') = annee_dem -1 then
      TobPal.PutValue('CATOT_ANT',TobPal.GetValue('CATOT_ANT')+ TOBT.GetValue('GZS_CATOT'))
    else
      TobPal.PutValue('CATOT',TobPal.GetValue('CATOT')+ TOBT.GetValue('GZS_CATOT'));
    end;
  end;
FiniMove;

// si cumul prograssif demandé, modification des totaux de la TOB
if (cumprog = True) then
  begin
  cum := 0;
  cum_ant := 0;
  for i_ind:=0 to TobTemp.Detail.count-1 do
    begin
    TobPal := TobTemp.Detail[i_ind];
    TobPal.PutValue('CATOT',TobPal.GetValue('CATOT')+ cum);
    TobPal.PutValue('CATOT_ANT',TobPal.GetValue('CATOT_ANT')+ cum_ant);
    cum := TobPal.GetValue('CATOT');
    cum_ant := TobPal.GetValue('CATOT_ANT');
    end;
  end;

end;

procedure TOF_GraphSynVte.regroupe_ligne_jour(TobSyn : TOB ; var TobTemp : TOB);
var TobPal,TobT : TOB;
    i_ind : integer;
begin
// chargement des cumul pour chaque jour
InitMove(TobSyn.Detail.count,'');
for i_ind:=0 to TobSyn.Detail.count-1 do
  begin
  TobT := TobSyn.Detail[i_ind];
  TobPal := TobTemp.FindFirst(['MOIS','JOUR'],[IntToStr(TobT.GetValue('GZS_MOIS')),IntToStr(TobT.GetValue('GZS_JOUR'))], False);
  MoveCur(False);
  if TobPal = nil then
    begin
    TobPal := TOB.Create('edtsyn',TobTemp,-1);
    TobPal.AddChampSup('MOIS',False);
    TobPal.AddChampSup('JOUR',False);
    TobPal.AddChampSup('CATOT',False);
    TobPal.AddChampSup('CATOT_ANT',False);
    TobPal.InitValeurs;
    TobPal.PutValue('MOIS',TobT.GetValue('GZS_MOIS'));
    TobPal.PutValue('JOUR',TobT.GetValue('GZS_JOUR')) ;
    TobPal.PutValue('CATOT',0);
    TobPal.PutValue('CATOT_ANT',0);
    end;
  if TobT.GetValue('GZS_ANNEE') = annee_dem -1 then
    TobPal.PutValue('CATOT_ANT',TobPal.GetValue('CATOT_ANT')+ TOBT.GetValue('GZS_CATOT'))
  else
    TobPal.PutValue('CATOT',TobPal.GetValue('CATOT')+ TOBT.GetValue('GZS_CATOT'));
  end;
FiniMove;
end;

procedure TOF_GraphSynVte.Cumul_ligne( var TobTemp : TOB);
var TobPal : TOB;
    i_ind : integer;
    cum, cum_ant : double;
begin
// si cumul prograssif demandé, modification des totaux de la TOB
cum := 0;
cum_ant := 0;
for i_ind:=0 to TobTemp.Detail.count-1 do
  begin
  TobPal := TobTemp.Detail[i_ind];
  TobPal.PutValue('CATOT',TobPal.GetValue('CATOT')+ cum);
  TobPal.PutValue('CATOT_ANT',TobPal.GetValue('CATOT_ANT')+ cum_ant);
  cum := TobPal.GetValue('CATOT');
  cum_ant := TobPal.GetValue('CATOT_ANT');
  end;
end;

Procedure TOF_GraphSynVte_SauveGraph (parms:array of variant; nb: integer) ;
var F : TFGRS1;
begin
F := TFGRS1(Longint (Parms[0]));
if ((F.Name) = 'GCGRAPHSYNM') or ((F.Name) = 'GCGRAPHSYNA') then
   begin
   if TCheckBox(F.FindComponent('COMPAR')).Checked then F.InitGR('T'+F.Name+'COMP')
   else F.InitGR('T'+F.Name);
   end
else if ((F.Name) = 'GCGRAPHVIEPROD') then
   begin
   if TCheckBox(F.FindComponent('ENTSTOCK')).Checked then F.InitGR('T'+F.Name+'ESTK')
   else F.InitGR('T'+F.Name);
   end;

end;

Initialization
registerclasses([TOF_GraphSynVte]);
RegisterAglProc ('SauveGraph', True , 1, TOF_GraphSynVte_SauveGraph);
end.
