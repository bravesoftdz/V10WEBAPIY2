unit UTofGraphVieProd;

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
  UTOB, HStatus, HPanel, M3FP, Hent1, EntGC, GraphUtil, Series, UtilArticle,UtilGC;

Type
     TOF_GraphVieProd = Class (TOF)
        procedure OnArgument (Argument : string); override;
        procedure OnUpdate; override;
        procedure regroupe_ligne(TobSyn : TOB; var TobTemp : TOB; Echelle : string);
        procedure RempliTOBVteSto(TobTemp : TOB ; var TobSyn : TOB);
        procedure ChargeGraph (Echelle : string);
        procedure OnClose  ; override ;
     private
        stcriteres, codeart : string;
        Ent_stock, progressif, choixFamille1, choixFamille2, choixFamille3,choixFamille : boolean;
        DateDeb, DateFin : TDateTime ;
        TobSyn : TOB;
        Annee, Mois, Jour, Semaine : Word;
     end;

implementation

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_GraphVieProd.OnArgument (Argument : string);
var F : TForm;
    iCol : integer ;
    THLIB : THLabel ;
begin
inherited;
F := TForm (Ecran);
for iCol:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+InttoStr(iCol),Ecran);
//THValComboBox(F.FindComponent ('ETABLISSEMENT')).Value := VH_GC.GCDepotDefaut;
TToolbarButton97(F.FindComponent('bAffGraph')).Down:=True;
THGrid(F.FindComponent('FListe')).Visible:=False;
end;

procedure TOF_GraphVieProd.OnUpdate;
Var F : TFGRS1;
    bFirstCrit : boolean ;
    stWhere, stnature : string;
    QVte, QSto, Qdis   : TQuery;
    TobVte,TobSto, TobT,qte_attendu : TOB;
    Edit : THEdit ;
    Edit2 : THValcombobox ;
    Edit3 : THRadioGroup ;
    ctrl, ctrlName, signe, Echelle, stTmp,stCrit : string;
    codeart_sv, codedep_sv, familleNiv1_sv, familleNiv2_sv, familleNiv3_sv : string;
     Stock_art : double;
    i_ind, i_ind2 : integer;
begin
inherited;
F := TFGRS1(Ecran);
Ent_stock := False;
Progressif := False;
if TcheckBox(TFGRS1(F).FindComponent('ENTSTOCK')).State = cbChecked then
   Ent_stock := True;
if TcheckBox(TFGRS1(F).FindComponent('PROGRES')).State = cbChecked then
   Progressif := True;
Edit3 := THRadioGroup(TFGRS1(F).FindComponent('EDTECHELLE'));
Echelle := Edit3.Values[Edit3.ItemIndex];
// Il faut détruire "TobSyn" si elle existe déjà, je le fait ici car on en a besoin dans
// ChargeGraph et il peut être appelé à tout moment, quand on passe dans ChangeAffichage
if TobSyn<>nil then
   begin
   TobSyn.free;
   TobSyn:=nil;
   end;
// modification pour que l'on puisse choisir plusieurs établissements dans un MultiComboBox
ctrl:='GL_ETABLISSEMENT';
 stCrit:=GetControlText('DEPOT');
if (stCrit<>'') and (stCrit<>'<<Tous>>') then
    BEGIN
    bFirstCrit:=True ;
    if stWhere<>'' then stWhere:=stWhere+' AND ' ;
    stWhere:=stWhere+ctrl+' in (' ;
    repeat
        stTmp:=ReadTokenSt(stCrit) ;
        if bFirstCrit then bFirstCrit:=False else stWhere:=stWhere+',' ;
        stWhere:=stWhere+'"'+stTmp+'"' ;
    until stCrit='' ;
    stWhere:=stWhere+')' ;
    END ;

// Ajout de la sélection suivant les familles et sous familles d'un article
choixFamille1 := False;
choixFamille2 := False;
choixFamille3 := False;
for i_ind:=1 to 3 do
   begin
   ctrl:='GL_FAMILLENIV'+InttoStr(i_ind);
   ctrlName:='GA_FAMILLENIV'+InttoStr(i_ind); signe:='=' ;
   Edit2:=THValComboBox(TFGRS1(F).FindComponent(ctrlName)) ;
   if (Edit2<>nil) and (Edit2.Value<>'') and (Edit2.Value<>'<<Tous>>') then
     begin
     if i_ind=1 then choixFamille1 := True;
     if i_ind=2 then choixFamille2 := True;
     if i_ind=3 then choixFamille3 := True;
     if stWhere<>'' then stWhere:=stWhere+' AND ' ;
     stWhere:=stWhere+ctrl+signe+'"'+Edit2.Value+'"' ;
     end ;
   end;
   

ctrl:='GL_ARTICLE';
ctrlName:='REFERENCE'; signe:=' like ' ;
Edit:=THEdit(TFGRS1(F).FindComponent(ctrlName)) ;
codeart := Edit.text;
if (Edit.text = '') and (choixFamille1 = false) and (choixFamille2 = false) and (choixFamille3 = false) then
   begin
   HShowMessage('0;Code article ou famille obligatoire;;E;O;O;O;','','') ;
   exit;
   end;
if ((choixFamille1 = False) and (choixFamille2 = False) and (choixFamille3 = False)) or (Edit.text <> '')then
   begin
   if stWhere<>'' then stWhere:=stWhere+' AND ' ;
   stWhere:=stWhere+ctrl+signe+'"'+codeart+'%"' ;
   end;


Edit:=THEdit(TFGRS1(F).FindComponent('DATEPIECE_DEB')) ;
DateDeb:= StrToDate(Edit.Text) ;
ctrl:='GL_DATEPIECE';
Edit:=THEdit(TFGRS1(F).FindComponent('DATEPIECE_FIN')) ;
DateFin:= StrToDate(Edit.Text) ;
if stWhere<>'' then stWhere:=stWhere+' AND ' ;
stWhere:=stWhere+ctrl+' >= "'+USDateTime(DateDeb)+'" AND ' ;
stWhere:=stWhere+ctrl+' <= "'+USDateTime(DateFin)+'"' ;

stnature := 'GL_NATUREPIECEG = "BLF" OR GL_NATUREPIECEG = "FFO"';
if (Ent_stock = True) then
  stnature := stnature +' OR GL_NATUREPIECEG = "TRE" OR GL_NATUREPIECEG = "TEM"';

QSto := OpenSQL('SELECT GL_ETABLISSEMENT, GL_CODEARTICLE, GL_FAMILLENIV1, GL_FAMILLENIV2,'+
' GL_FAMILLENIV3, GL_NATUREPIECEG, GL_DATEPIECE, GL_QTEFACT FROM LIGNE '+
'WHERE ('+stnature+') AND GL_INDICEG = 0 AND '+
stWhere+' ORDER BY GL_ETABLISSEMENT, GL_CODEARTICLE, GL_DATEPIECE ', True);

if QSto.Eof then
   begin
   HShowMessage('0;Pas de Données sur cette période;;E;O;O;O;','','') ;
   Ferme(QSto);
   exit;
   end;

TobSto:=TOB.Create('',nil,-1);
TobSto.LoadDetailDB('','','',QSto,false);
Ferme(QSto);

// la date de début de consultation pour recalcul du stock à la date J
codeart_sv := '';
codedep_sv := '';
familleNiv1_sv := '';
familleNiv2_sv := '';
familleNiv3_sv := '';
Stock_art := 0;
qte_attendu := nil;
for i_ind:=0 to TobSto.Detail.count-1 do
   begin
   TobT := TobSto.Detail[i_ind];
// Suivant si on a choisi dans le mul un article ou une famille ; sous-famille, j'appelle
// la fonction du calcul du stock avec les bons paramètres correspondant à ce choix
   if (codeart<>'' )then
      begin
      if not (codeart_sv = TobT.GetValue('GL_CODEARTICLE'))
      or not (codedep_sv = TobT.GetValue('GL_ETABLISSEMENT')) then
         begin
         RecalcStockMois(TobT.GetValue('GL_ETABLISSEMENT'), TobT.GetValue('GL_CODEARTICLE'), '', '', '', datedeb, Stock_art, qte_attendu, 'STO');
         end;
      end
   else if (choixFamille1 = True) and (choixFamille2 = False) and (choixFamille3 = False) then
      begin
      if not (codedep_sv = TobT.GetValue('GL_ETABLISSEMENT'))
      or not (familleNiv1_sv = TobT.GetValue('GL_FAMILLENIV1')) then
         begin
         RecalcStockMois(TobT.GetValue('GL_ETABLISSEMENT'), '', TobT.GetValue('GL_FAMILLENIV1'), '', '', datedeb, Stock_art, qte_attendu, 'STO');
         end;
      end
   else  if (choixFamille1 = True) and (choixFamille2 = True) and (choixFamille3 = False) then
      begin
      if not (codedep_sv = TobT.GetValue('GL_ETABLISSEMENT'))
      or not (familleNiv1_sv = TobT.GetValue('GL_FAMILLENIV1'))
      or not (familleNiv2_sv = TobT.GetValue('GL_FAMILLENIV2')) then
         begin
         RecalcStockMois(TobT.GetValue('GL_ETABLISSEMENT'), '',TobT.GetValue('GL_FAMILLENIV1'),TobT.GetValue('GL_FAMILLENIV2'),'', datedeb, Stock_art, qte_attendu, 'STO');
         end;
      end
   else  if (choixFamille1 = True) and (choixFamille2 = False) and (choixFamille3 = True) then
      begin
      if not (codedep_sv = TobT.GetValue('GL_ETABLISSEMENT'))
      or not (familleNiv1_sv = TobT.GetValue('GL_FAMILLENIV1'))
      or not (familleNiv3_sv = TobT.GetValue('GL_FAMILLENIV3')) then
         begin
         RecalcStockMois(TobT.GetValue('GL_ETABLISSEMENT'), '',TobT.GetValue('GL_FAMILLENIV1'),'',TobT.GetValue('GL_FAMILLENIV3'), datedeb, Stock_art, qte_attendu, 'STO');
         end;
      end
   else  if (choixFamille1 = True) and (choixFamille2 = True) and (choixFamille3 = True) then
      begin
      if not (codedep_sv = TobT.GetValue('GL_ETABLISSEMENT'))
      or not (familleNiv1_sv = TobT.GetValue('GL_FAMILLENIV1'))
      or not (familleNiv2_sv = TobT.GetValue('GL_FAMILLENIV2'))
      or not (familleNiv3_sv = TobT.GetValue('GL_FAMILLENIV3')) then
         begin
         RecalcStockMois(TobT.GetValue('GL_ETABLISSEMENT'), '',TobT.GetValue('GL_FAMILLENIV1'),TobT.GetValue('GL_FAMILLENIV2'),TobT.GetValue('GL_FAMILLENIV3'), datedeb, Stock_art, qte_attendu, 'STO');
         end;
      end
   else if (choixFamille1 = False) and (choixFamille2 = True) and (choixFamille3 = True) then
      begin
      if not (codedep_sv = TobT.GetValue('GL_ETABLISSEMENT'))
      or not (familleNiv2_sv = TobT.GetValue('GL_FAMILLENIV2'))
      or not (familleNiv3_sv = TobT.GetValue('GL_FAMILLENIV3')) then
         begin
         RecalcStockMois(TobT.GetValue('GL_ETABLISSEMENT'), '','',TobT.GetValue('GL_FAMILLENIV2'),TobT.GetValue('GL_FAMILLENIV3'), datedeb, Stock_art, qte_attendu, 'STO');
         end;
      end
   else if (choixFamille1 = False) and (choixFamille2 = True) and (choixFamille3 = False) then
      begin
      if not (codedep_sv = TobT.GetValue('GL_ETABLISSEMENT'))
      or not (familleNiv2_sv = TobT.GetValue('GL_FAMILLENIV2')) then
         begin
         RecalcStockMois(TobT.GetValue('GL_ETABLISSEMENT'), '','',TobT.GetValue('GL_FAMILLENIV2'),'', datedeb, Stock_art, qte_attendu, 'STO');
         end;
      end
   else if (choixFamille1 = False) and (choixFamille2 = False) and (choixFamille3 = True) then
      begin
      if not (codedep_sv = TobT.GetValue('GL_ETABLISSEMENT'))
      or not (familleNiv3_sv = TobT.GetValue('GL_FAMILLENIV3')) then
         begin
         RecalcStockMois(TobT.GetValue('GL_ETABLISSEMENT'), '','','',TobT.GetValue('GL_FAMILLENIV3'), datedeb, Stock_art, qte_attendu, 'STO');
         end;
      end;

   TobT.AddChampSup('STOCK',False);
   TobT.PutValue('STOCK',Stock_art);
   IF Not VarIsNull(TobT.GetValue('GL_CODEARTICLE')) then codeart_sv := TobT.GetValue('GL_CODEARTICLE') else codeart_sv :='';
   codedep_sv := TobT.GetValue('GL_ETABLISSEMENT');
   IF Not VarIsNull(TobT.GetValue('GL_FAMILLENIV1')) then familleNiv1_sv := TobT.GetValue('GL_FAMILLENIV1') else familleNiv1_sv :='';
   IF Not VarIsNull(TobT.GetValue('GL_FAMILLENIV2')) then familleNiv2_sv := TobT.GetValue('GL_FAMILLENIV2') else familleNiv2_sv :='';
   IF Not VarIsNull(TobT.GetValue('GL_FAMILLENIV3')) then familleNiv3_sv := TobT.GetValue('GL_FAMILLENIV3') else familleNiv3_sv :='';
   end;

TobSyn := TOB.Create('VTESTO',nil,-1);
RempliTOBVteSto(TobSto, TobSyn);

TobSto.Free;
ChargeGraph (Echelle);
end;

procedure TOF_GraphVieProd.OnClose  ;
begin
if TobSyn<>nil then
TobSyn.Free;  
end;

procedure TOF_GraphVieProd.RempliTOBVteSto(TobTemp : TOB ; var TobSyn : TOB);
var TobPal, TOBT : TOB;
    An_deb, Mo_deb, Jo_deb, Sem_deb : Word;
    i_ind : integer;
    LaDate : TDateTime ;
begin
DecodeDate(DateDeb, An_deb, Mo_deb, Jo_deb);
Sem_deb := NumSemaine (DateDeb) ;
InitMove(TobTemp.Detail.count,'');
for i_ind:=0 to TobTemp.Detail.count-1 do
  begin
  TobT := TobTemp.Detail[i_ind];
  Ladate := TobT.GetValue('GL_DATEPIECE');
  DecodeDate(LaDate, Annee, Mois, Jour);
  Semaine := NumSemaine (Ladate) ;

  if (Ladate < DateDeb) then
    begin
    Annee := An_deb;
    Mois := Mo_deb;
    Semaine := Sem_deb;
    end;

  TobPal := TobSyn.FindFirst(['ETABLISSEMENT','ANNEE','MOIS','SEMAINE'],
  [TobT.GetValue('GL_ETABLISSEMENT'),IntToStr(Integer(Annee)),IntToStr(Integer(mois)),IntToStr(Integer(Semaine))], False);

  MoveCur(False);
  if TobPal = nil then
     begin
     TobPal := TOB.Create('VteSto',TobSyn,-1);
     TobPal.AddChampSup('ETABLISSEMENT',False);
     TobPal.AddChampSup('ANNEE',False);
     TobPal.AddChampSup('MOIS',False);
     TobPal.AddChampSup('SEMAINE',False);
     TobPal.AddChampSup('LIBMOIS',False);
     TobPal.AddChampSup('QTEVTE',False);
     TobPal.AddChampSup('QTESTO',False);
     TobPal.AddChampSup('QTETRE',False);
     TobPal.AddChampSup('QTEREC',False);
     TobPal.InitValeurs;
     TobPal.PutValue('ETABLISSEMENT',TOBT.GetValue('GL_ETABLISSEMENT'));
     TobPal.PutValue('ANNEE',annee);
     TobPal.PutValue('MOIS',mois);
     TobPal.PutValue('SEMAINE',Semaine);
     TobPal.PutValue('LIBMOIS',ShortMonthNames[mois]);
     TobPal.PutValue('QTEVTE',0);
     TobPal.PutValue('QTESTO',TOBT.GetValue('STOCK'));
     TobPal.PutValue('QTETRE',0);
     TobPal.PutValue('QTEREC',0);
     end;
  if (TOBT.GetValue('GL_NATUREPIECEG') = 'FFO') or (TOBT.GetValue('GL_NATUREPIECEG') = 'TEM') then
    TobPal.PutValue('QTEVTE',TobPal.GetValue('QTEVTE')+TOBT.GetValue('GL_QTEFACT'))
  else
    if (TOBT.GetValue('GL_NATUREPIECEG') = 'TRE') then
      TobPal.PutValue('QTETRE',TobPal.GetValue('QTETRE')+TOBT.GetValue('GL_QTEFACT'))
    else
      TobPal.PutValue('QTEREC',TobPal.GetValue('QTEREC')+TOBT.GetValue('GL_QTEFACT'));
  end;
FiniMove;
end;

procedure TOF_GraphVieProd.ChargeGraph (Echelle : string);
var Tobtmp : TOB;
    F : TFGRS1;
    sttitre, stColonnes, stChampTitre : string;
    tstTitres : Tstrings;
    stTitresCol1, stTitresCol2, stColGraph1, stColGraph2 : string;
    stWhere : string;
    Serie : TChartSeries ;
begin
stCriteres := '';
stColGraph1 := '';
stColGraph2 := '';
stTitresCol1 := '';
stTitresCol2 := '';
sttitre := '';

tstTitres := TStringList.Create;
F := TFGRS1(Ecran);

Tobtmp:=TOB.Create('EDTEVTSTO',nil,-1);
regroupe_ligne(TobSyn, Tobtmp, Echelle);
if (Echelle = 'A') then
  begin
  if (Ent_stock = True) then
    begin
    stColonnes := 'ANNEE;QTESTO;QTEREC;QTETRE;QTEVTE';
    stTitre := 'Année;Stock Initial;Réceptions;Transferts reçus;Ventes & Transferts émis';            // légende des colonnes
    end
  else
    begin
    stColonnes := 'ANNEE;QTESTO;QTEREC;QTEVTE';
    stTitre := 'Année;Stock Initial;Réceptions;Ventes ';           // légende des colonnes
    end;
  end;
if (Echelle = 'M') then
  begin
  if (Ent_stock = True) then
    begin
    stColonnes := 'LIBMOIS;QTESTO;QTEREC;QTETRE;QTEVTE';
    stTitre := 'Mois;Stock Initial;Réceptions;Transferts reçus;Ventes & Transferts émis';            // légende des colonnes
    end
  else
    begin
    stColonnes := 'LIBMOIS;QTESTO;QTEREC;QTEVTE';
    stTitre := 'Mois;Stock Initial;Réceptions;Ventes ';           // légende des colonnes
    end;
  end;
if (Echelle = 'S') then
  begin
  if (Ent_stock = True) then
    begin
    stColonnes := 'SEMAINE;QTESTO;QTEREC;QTETRE;QTEVTE';
    stTitre := 'Semaine;Stock Initial;Réceptions;Transferts reçus;Ventes & Transferts émis';
    end           // légende des colonnes
  else
    begin
    stColonnes := 'SEMAINE;QTESTO;QTEREC;QTEVTE';
    stTitre := 'Semaine;Stock Initial;Réceptions;Ventes';            // légende des colonnes
    end;
  end;
//tstTitres.Add ('Courbe de vie de l''article '+codeart);
tstTitres.Add ('Courbe de vie ');
tstTitres.Add ('pour la période du '+DateToStr(datedeb)+' au '+DateToStr(datefin));
if (Ent_stock = True) then
  stColGraph1 := 'QTESTO;QTEREC;QTETRE;QTEVTE'
else
  stColGraph1 := 'QTESTO;QTEREC;QTEVTE';
stChampTitre := stTitre;

LanceGraph (F, Tobtmp, '', stColonnes,  '', stTitre ,
            stColGraph1, stColGraph2, tstTitres, nil, TBarSeries, stChampTitre, False);
F.FChart1.SeriesList.series[0].Marks.Visible := False;
F.FChart1.SeriesList.series[1].Marks.Visible := False;
F.FChart1.SeriesList.series[2].Marks.Visible := False;
if (Ent_stock = True) then
  F.FChart1.SeriesList.series[3].Marks.Visible := False;

F.FChart1.SeriesList.series[0].SeriesColor := clLime;
if (Ent_stock = True) then
  begin
  F.FChart1.SeriesList.series[1].SeriesColor := clBlue;
  F.FChart1.SeriesList.series[2].SeriesColor := clYellow;
  F.FChart1.SeriesList.series[3].SeriesColor := clRed;
  ( F.FChart1.SeriesList.series[1] as TBarSeries ).Multibar := mbStacked ;
  Serie:=F.FChart1.SeriesList.series[3] ;
  ChangeSeriesType(serie,TLineSeries);
  end
else
  begin
  ( F.FChart1.SeriesList.series[1] as TBarSeries ).Multibar := mbStacked ;
  F.FChart1.SeriesList.series[1].SeriesColor := clBlue;
  F.FChart1.SeriesList.series[2].SeriesColor := clRed;
  Serie:=F.FChart1.SeriesList.series[2] ;
  ChangeSeriesType(serie,TLineSeries);
  end;
F.FChart1.LeftAxis.Title.caption := 'Quantité';
if (Echelle = 'A') then
  F.FChart1.BottomAxis.Title.caption := 'Année';
if (Echelle = 'M') then
  F.FChart1.BottomAxis.Title.caption := 'Mois';
if (Echelle = 'S') then
  F.FChart1.BottomAxis.Title.caption := 'Semaine';
Tobtmp.Free;
tstTitres.Free;
end;

procedure TOF_GraphVieProd.regroupe_ligne(TobSyn : TOB ; var TobTemp : TOB; Echelle : string);
var TobPal,TobT : TOB;
    i_ind, i : integer;
    testdate : TDateTime;
    cumrec, cumtre, cumvte, cumsto : double;
begin

// génération des cumuls suivant le critère demandé
if TobSyn <> nil then
   begin
   InitMove(TobSyn.Detail.count,'');
   TobSyn.Detail.Sort('ANNEE;MOIS;SEMAINE') ;
   testdate := datedeb;
   while testdate <= datefin do
      begin
      DecodeDate(testdate, annee, mois, Jour);
      Semaine := NumSemaine (testdate) ;
      if (Echelle = 'A') then
       TobPal := TobTemp.FindFirst(['ANNEE'],[IntToStr(annee)], False);
      if (Echelle = 'M') then
       TobPal := TobTemp.FindFirst(['ANNEE','MOIS'],[IntToStr(annee),IntToStr(mois)], False);
      if (Echelle = 'S') then
       TobPal := TobTemp.FindFirst(['ANNEE','SEMAINE'],[IntToStr(annee),IntToStr(semaine)], False);
      MoveCur(False);
      if TobPal = nil then
         begin
         TobPal := TOB.Create('EdtVteSto',TobTemp,-1);
         TobPal.AddChampSup('ANNEE',False);
         TobPal.AddChampSup('MOIS',False);
         TobPal.AddChampSup('SEMAINE',False);
         TobPal.AddChampSup('LIBMOIS',False);
         TobPal.AddChampSup('QTEVTE',False);
         TobPal.AddChampSup('QTESTO',False);
         TobPal.AddChampSup('QTETRE',False);
         TobPal.AddChampSup('QTEREC',False);
         TobPal.InitValeurs;
         TobPal.PutValue('ANNEE',annee);
         if (Echelle = 'M') then
           TobPal.PutValue('MOIS',mois);
         if (Echelle = 'S') then
           TobPal.PutValue('SEMAINE',semaine);
         TobPal.PutValue('LIBMOIS',ShortMonthNames[mois]);
         TobPal.PutValue('QTEVTE',0);
         TobPal.PutValue('QTESTO',0);
         TobPal.PutValue('QTETRE',0);
         TobPal.PutValue('QTEREC',0);
         end;
      testdate := PlusDate(testdate,1,'J');
      end;
   for i_ind:=0 to TobSyn.Detail.count-1 do
      begin
      TobT := TobSyn.Detail[i_ind];
      if (Echelle = 'A') then
       TobPal := TobTemp.FindFirst(['ANNEE'],[IntToStr(TOBT.GetValue('ANNEE'))], False);
      if (Echelle = 'M') then
       TobPal := TobTemp.FindFirst(['ANNEE','MOIS'],[IntToStr(TOBT.GetValue('ANNEE')),IntToStr(TOBT.GetValue('MOIS'))], False);
      if (Echelle = 'S') then
       TobPal := TobTemp.FindFirst(['ANNEE','SEMAINE'],[IntToStr(TOBT.GetValue('ANNEE')),IntToStr(TOBT.GetValue('SEMAINE'))], False);
      MoveCur(False);
      if TobPal = nil then
         begin
         TobPal := TOB.Create('EdtVteSto',TobTemp,-1);
         TobPal.AddChampSup('ANNEE',False);
         TobPal.AddChampSup('MOIS',False);
         TobPal.AddChampSup('SEMAINE',False);
         TobPal.AddChampSup('LIBMOIS',False);
         TobPal.AddChampSup('QTEVTE',False);
         TobPal.AddChampSup('QTESTO',False);
         TobPal.AddChampSup('QTETRE',False);
         TobPal.AddChampSup('QTEREC',False);
         TobPal.InitValeurs;
         TobPal.PutValue('ANNEE',TOBT.GetValue('ANNEE'));
         if (Echelle = 'M') then
          TobPal.PutValue('MOIS',TOBT.GetValue('MOIS'));
         if (Echelle = 'S') then
          TobPal.PutValue('SEMAINE',TOBT.GetValue('SEMAINE'));
         TobPal.PutValue('LIBMOIS',TOBT.GetValue('LIBMOIS'));
         TobPal.PutValue('QTEVTE',0);
         TobPal.PutValue('QTESTO',0);
         TobPal.PutValue('QTETRE',0);
         TobPal.PutValue('QTEREC',0);
         end;
      TobPal.PutValue('QTEVTE',TobPal.GetValue('QTEVTE')+TOBT.GetValue('QTEVTE'));
      TobPal.PutValue('QTESTO',TobT.GetValue('QTESTO'));
      TobPal.PutValue('QTETRE',TobPal.GetValue('QTETRE')+TOBT.GetValue('QTETRE'));
      TobPal.PutValue('QTEREC',TobPal.GetValue('QTEREC')+TOBT.GetValue('QTEREC'));
      end;

   cumsto := 0;
   cumrec := 0;
   cumtre := 0;
   cumvte := 0;
   for i_ind:=0 to TobTemp.Detail.count-1 do
      begin
      TobT := TobTemp.Detail[i_ind];
      if TobT.GetValue('QTESTO') <> 0 then  cumsto := TobT.GetValue('QTESTO');
      TobT.PutValue('QTESTO',cumsto);
      if (progressif = True) then
         begin
         cumvte := cumvte + TobT.GetValue('QTEVTE');
         cumtre := cumtre + TobT.GetValue('QTETRE');
         cumrec := cumrec + TobT.GetValue('QTEREC');
         TobT.PutValue('QTEVTE',cumvte);
         TobT.PutValue('QTETRE',cumtre);
         TobT.PutValue('QTEREC',cumrec);
         end;
      end;
   FiniMove;
   end
else  HShowMessage('0;Pas de Données sur cette période;;E;O;O;O;','','') ;
end;

Procedure TOF_GraphVieProd_ChangeAffichage (parms:array of variant; nb: integer );
var F : TFGRS1;
     TOTOF : TOF ;
begin
F:=TFGRS1(Longint(Parms[0])) ;
if (F.Name <> 'GCGRAPHVIEPROD') then exit;
if (F is TFGRS1) then TOTOF:=TFGRS1(F).Latof else exit;
if (TOTOF is TOF_GraphVieProd) then
  TOF_GraphVieProd(TOTOF).ChargeGraph (Parms[1]) else exit;
end;

procedure InitTOFGraphVieProd ();
begin
RegisterAglProc('GraphVieProd_ChangeAffichage', True , 1, TOF_GraphVieProd_ChangeAffichage);
end;

Initialization
registerclasses([TOF_GraphVieProd]);
InitTOFGraphVieProd();
end.
