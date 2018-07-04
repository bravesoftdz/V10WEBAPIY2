unit UTofGraphPalm;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Db, DBTables, Hqry, HSysMenu, hmsgbox, Menus, ComCtrls, TeeProcs,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      Fiche, Fe_Main,
{$ENDIF}
  GRS1, TeEngine, Chart, Grids, Hctrls, ExtCtrls, StdCtrls, HTB97, UIUtil, UTOF,
  UTOB, HStatus, UtilGC, HPanel, M3FP, Hent1, EntGC, GraphUtil, Series, UtilArticle,
  UTofEditStat, HDimension;

Const
	Titre: array[1..3] of string 	= (
          {1}         'ARTICLE',
          {2}         'COMMERCIAL',
          {3}         ''
                 );
    SecondTitre : array [1..3] of string = (
          {1}         'par article ',
          {2}         'par représentant ',
          {3}         ''
                 );

Type
     TOF_GraphPalm = Class (TOF)
        procedure OnArgument (stArgument : string); override;
        procedure OnUpdate; override;

        function ChoixRequete (iChoix : integer) : string;
        procedure CpteLigne(TobPal : TOB);
        procedure RecupCritere(Choix : integer);
        procedure RempliTOBPalmares(TobT : TOB; var TobPalmares : TOB);
        procedure ZoomGraphe (Sender: TCustomChart; Series: TChartSeries;
                              ValueIndex: Integer; Button: TMouseButton;
                              Shift: TShiftState; X, Y: Integer);
     private
        stcriteres : string;
        stOrderBy : string;
        stNatureauxi : string;
        i_cle : integer;
     end;

implementation

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
procedure TOF_GraphPalm.OnArgument (stArgument : string);
var Iind : integer;
    stPlus, stDef : string ;
begin
inherited;
if stArgument = 'VEN' then
    begin
    stNatureauxi := 'CLI';
{$IFDEF NOMADE}
    stPlus := 'AND GPP_VENTEACHAT="VEN' +
              '" AND (GPP_NATUREPIECEG IN ("CC","DE") OR GPP_TRSFVENTE="X")' ;
{$ELSE}
    stPlus := 'AND GPP_VENTEACHAT="VEN"' ;
{$ENDIF}
    SetControlProperty ( 'NATUREPIECEG', 'Plus', stPlus ) ;
    Ecran.Caption := TraduireMemoire ( 'Palmarès des ventes' ) ;
    TRadioButton(GetControl('RTIERS')).Caption := TraduireMemoire( 'Client' ) ;
    TRadioButton(GetControl('RMARGE')).Enabled := True;
    Titre[3] := 'CLIENT';
    SecondTitre[3] := traduireMemoire ( 'par client' ) ;
{$IFDEF NOMADE}
    stDef := 'CC' ;
{$ELSE}
    if not(ctxAffaire in V_PGI.PGIContexte)  then
    Begin
      FactureAvoir_AddItem(GetControl('NATUREPIECEG'),'VEN');
      stDef := 'ZZV' ;
    End;
 {$ENDIF}
    SetControlText ( 'NATUREPIECEG', stDef ) ;
    //mcd 27/03/02 *** Blocages des natures de pièces autorisées en affaires
    if (ctxAffaire in V_PGI.PGIContexte)  then
       begin
       SetCOntrolVisible('DEPOT',False);
       SetCOntrolVisible('TDEPOT',False);
       if (ctxScot in V_PGI.PGIContexte)  then
       		THValComboBox(GetControl('NATUREPIECEG')).Plus :=  'AND ((GPP_NATUREPIECEG="APR") or (GPP_NATUREPIECEG="FPR") or (GPP_NATUREPIECEG="FAC")or (GPP_NATUREPIECEG="AVC") or (GPP_NATUREPIECEG="FRE"))';
          // mcd 10/10/02 obligation de mettre après car condition plus changée
        Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FAC');
        if Iind<0 then Iind:=0;
        THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Avoir clients + Fac. Reprise');
        THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ1');
        SetControlText('NATUREPIECEG','ZZ1');

        Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FRE');
        if Iind<0 then Iind:=0;
        THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Fac. Reprise');
        THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ2');
        SetControlText('NATUREPIECEG','ZZ2');

       end;
    // **** Fin affaire ***
    end else
    begin
    stNatureauxi := 'FOU';
{$IFDEF NOMADE}
    stPlus := 'AND GPP_VENTEACHAT="ACH' +
              '" AND (GPP_NATUREPIECEG="CF" OR GPP_TRSFACHAT="X")' ;
{$ELSE}
    stPlus := 'AND GPP_VENTEACHAT="ACH"' ;
{$ENDIF}
    SetControlProperty ( 'NATUREPIECEG', 'Plus', stPlus ) ;
    Ecran.Caption := TraduireMemoire ( 'Palmarès des achats' ) ;
    TRadioButton(GetControl('RTIERS')).Caption := TraduireMemoire ( 'Fournisseur' ) ;
    TRadioButton(GetControl('RMARGE')).Enabled := False;
    Titre[3] := 'FOURNISSEUR';
    SecondTitre[3] := TraduireMemoire ( 'par fournisseur' ) ;
{$IFDEF NOMADE}
    stDef := 'CF' ;
{$ELSE}
    FactureAvoir_AddItem(GetControl('NATUREPIECEG'),'ACH');
    stDef := 'BLF' ;
{$ENDIF}
    SetControlText ( 'NATUREPIECEG', stDef ) ;
    SetControlProperty('RGVENTES','Caption', TraduireMemoire ( 'Achats' ) ) ;
    SetControlProperty('RMEILLEURES','Caption',TraduireMemoire ( 'Les meilleurs' ) ) ;
    SetControlProperty('RPETITES','Caption',TraduireMemoire ( 'Les plus petits' ) ) ;
    //*** Blocages des natures de pièces autorisées en affaires
    if (ctxAffaire in V_PGI.PGIContexte)  then
       begin
       SetCOntrolVisible('DEPOT',False);
       SetCOntrolVisible('TDEPOT',False);
       THValComboBox(GetControl('NATUREPIECEG')).Plus :='AND (GPP_VENTEACHAT="ACH") '+  AfPlusNatureAchat;
       SetcontrolTExt('NATUREPIECEG','FF');
       end;
    // **** Fin affaire ***
    end;
TToolbarButton97(GetControl('bAffGraph')).Down:=True;
THGrid(GetControl('FListe')).Visible:=False;
end;

procedure TOF_GraphPalm.OnUpdate;
Var F : TFGRS1;
    stRequete, stSecondTitre, stColonnes, stChampTitre : string;
    tstTitres : Tstrings;
    TQSql : TQuery;
    iPalm,i_ind :integer;
    TobTemp, TobPalmares,TobP : TOB;
    stTitresCol1, stTitresCol2, stColGraph1, stColGraph2 : string;
begin
inherited;
F := TFGRS1(Ecran);
stCriteres := '';
stColGraph1 := '';
stColGraph2 := '';
stTitresCol1 := '';
stTitresCol2 := '';
stOrderBy := '';
i_cle := 0;
if THValComboBox(GetControl('NATUREPIECEG')).Value = '' then
    begin
    PGIBox('Vous devez renseigner une nature de pièce','Palmarès');
    exit;
    end;

stSecondTitre := '';
if (stNatureauxi = 'CLI') then
    begin
    if TRadioButton(GetControl('RMEILLEURES')).Checked = true then
        begin
        if StrToInt(GetControlText('SELECTION')) > 1 then stSecondTitre := 'Les ' + GetControlText('SELECTION') + ' meilleures ventes '
        else stSecondTitre := 'La meilleure vente ';
        end else
        begin
        if StrToInt(GetControlText('SELECTION')) > 1 then stSecondTitre := 'Les ' + GetControlText('SELECTION') + ' plus petites ventes '
        else stSecondTitre := 'La plus petite vente ';
        end;
    end else
    begin
    if TRadioButton(GetControl('RMEILLEURES')).Checked = true then
        begin
        if StrToInt(GetControlText('SELECTION')) > 1 then stSecondTitre := 'Les ' + GetControlText('SELECTION') + ' meilleurs achats '
        else stSecondTitre := 'Le meilleur achat ';
        end else
        begin
        if StrToInt(GetControlText('SELECTION')) > 1 then stSecondTitre := 'Les ' + GetControlText('SELECTION') + ' plus petits achats '
        else stSecondTitre := 'Le plus petit achat ';
        end;
    end;

tstTitres := TStringList.Create;
if TRadioButton(GetControl('RARTICLE')).Checked = True then iPalm := 1
   else if TRadioButton(GetControl('RCOMMERCIAL')).Checked = True then iPalm := 2
        else iPalm := 3;
stSecondTitre := stSecondTitre + SecondTitre[iPalm];

stRequete := ChoixRequete(iPalm);
TQSql := OpenSQL(stRequete,True);
TobTemp := TOB.Create('',nil,-1);

TobPalmares := TOB.Create('GCTMPPAL',nil,-1);
TQSql.First;
while not TQSql.Eof do
    begin
    TobTemp.SelectDb ('', TQSql);
    RempliTOBPalmares(TobTemp, TobPalmares);
    TQSql.Next;
    i_cle := i_cle + 1;
    end;
TobTemp.Free;
Ferme(TQSql);

if TobPalmares.Detail.Count > 0 then
    begin
    CpteLigne (TobPalmares);
    end;

for i_ind := 0 to TobPalmares.Detail.Count -1 do
    begin
    TobP := TobPalmares.Detail[i_ind];
    TobP.PutValue('GZA_PALMARES',copy(TobP.GetValue('GZA_PALMARES'),1,18) + ' ' + TobP.GetValue('DIMENSIONS'));
    TobP.PutValue('GZA_LIBELLE',TobP.GetValue('GZA_LIBELLE') + ' ' + TobP.GetValue('DIMENSIONS'));
    end;

stColonnes := 'GZA_PALMARES;GZA_LIBELLE;';
if TRadioButton(GetControl('RCODE')).Checked = True then
    begin
    stChampTitre := 'GZA_PALMARES';
    end;
if TRadioButton(GetControl('RLIBELLE')).Checked = True then
    begin
    stChampTitre := 'GZA_LIBELLE';
    end;

if TRadioButton(GetControl('RQUANTITE')).Checked = True then
    begin
    stColonnes := stColonnes + 'GZA_QTEFACT';
    stColGraph1 := 'GZA_QTEFACT';
    stTitresCol1 := 'Quantité';
    stSecondTitre := stSecondTitre + 'par quantité';
    end else if TRadioButton(GetControl('RCA')).Checked = True then
    begin
    stColonnes := stColonnes + 'GZA_TOTALHT';
    stColGraph1 := 'GZA_TOTALHT';
    stTitresCol1 := 'Chiffre d''affaire';
    stSecondTitre := stSecondTitre + 'par chiffre d''affaire';
    end else
    begin
    stColonnes := stColonnes + 'GZA_MARGE';
    stColGraph1 := 'GZA_MARGE';
    stTitresCol1 := 'Marge';
    stSecondTitre := stSecondTitre + 'par marge';
    end;

tstTitres.Add ('Palmares '+ Titre[iPalm]);
tstTitres.Add (stSecondTitre);
LanceGraph (F, TobPalmares, '', stColonnes,  '', Titre[iPalm] + ';libellé;' + stTitresCol1,
            stColGraph1, stColGraph2, tstTitres, nil, TPieSeries, stChampTitre, False);

if TCheckBox(GetControl('RPOURCENTAGE')).Checked = True then
    begin
    F.FChart1.Legend.LegendStyle := lsValues;
    F.FChart1.Legend.TextStyle := ltsLeftPercent;
    end else
    begin
    F.FChart1.Legend.LegendStyle := lsValues;
    F.FChart1.Legend.TextStyle := ltsLeftValue;
    end;

F.FChart1.OnClickSeries:=ZoomGraphe;
TobPalmares.free;
tstTitres.Free;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function TOF_GraphPalm.ChoixRequete(iChoix : integer) : string;
var stChampsCalc, Reste : string;
    stDim, Palmares : string;
    stOrdre, stPlus, stMarge : string;
    NumOrdre : integer;
begin
NumOrdre := 0;
stPlus:=''; stDim := '';
if iChoix=3 then stPlus := 'TOTAL';

// Récupération des conditions du lanceur d'état et ajout du préfixe des champs,
//les tables n'étant pas les mêmes d'un Palmarès à l'autre
RecupCritere(iChoix);
if TRadioButton(GetControl('RMEILLEURES')).Checked = True then stOrdre := ' DESC'
else stOrdre := ' ASC';

//Tri sur le CA ou sur la Quantité
if TRadioButton(GetControl('RQUANTITE')).Checked = True then
   stOrderBy := ' ORDER BY ' + IntToStr(4 + NumOrdre*2) + stOrdre
   else if TRadioButton(GetControl('RCA')).Checked = True then
   stOrderBy := ' ORDER BY ' + IntToStr(3 + NumOrdre*2) + stOrdre
   else stOrderBy := ' ORDER BY ' + IntToStr(6 + NumOrdre*2) + stOrdre;

stMarge := 'GL_'+VH_GC.GCMargeArticle;
if (stMarge = 'GL_PMA') or (stMarge = 'GL_PMR') then stMarge := stMarge + 'P';

//  BBI :Reste:='GL_QTEFACT-GL_QTERESTE';
Reste:='GL_QTERESTE';

stChampsCalc := ', SUM((GL_TOTALHT*('+Reste+'))/GL_QTEFACT) AS TOTALHT, '
              +'SUM('+Reste+') AS QTEFACT';
if ichoix <> 3 then
stChampsCalc := stChampsCalc + ',SUM('+ stMarge +'*'+'('+Reste+')) AS TOTALHTACH';

//Suppresion des lignes de commentaire
stCriteres := stCriteres + ' AND GL_QTEFACT<>0 AND GL_TOTALHT<>0 AND GL_TYPELIGNE="ART" ';

//Choix de la requête suivant s'il s'agit d'un Palmarès article,
//représentant ou client
case iChoix of
   1 : begin
       if (ctxMode in V_PGI.PGIContexte) then
          begin
          Palmares := 'CODEARTICLE';
          stCriteres := stCriteres + ' AND GA_STATUTART<>"DIM" ';
          end else
          begin
          Palmares := 'ARTICLE';
          stDim := ',GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,'+
                 'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5 ';
          end;
       Result := 'SELECT GL_'+Palmares+' AS PALMARES, GA_LIBELLE AS LIBELLE' +
                 stChampsCalc + stDim + ',GA_STATUTART FROM LIGNE , ARTICLE ' +
                 'WHERE GL_'+Palmares+'=GA_'+Palmares+ stCriteres +
                 ' GROUP BY GL_'+Palmares+', GA_LIBELLE,GA_STATUTART' + stDim + stOrderBy;
       end;
   2 : Result := 'SELECT GL_REPRESENTANT AS PALMARES, GCL_LIBELLE AS LIBELLE' +
                 stChampsCalc + ' FROM LIGNE, COMMERCIAL ' +
                 'WHERE GL_REPRESENTANT=GCL_COMMERCIAL' + stCriteres +
                 ' GROUP BY GL_REPRESENTANT, GCL_LIBELLE ' + stOrderBy;

   3 : Result := 'SELECT GL_TIERS AS PALMARES, T_LIBELLE AS LIBELLE' +
                 stChampsCalc + ' FROM LIGNE, TIERS ' +
                 'WHERE GL_TIERS=T_TIERS AND T_NATUREAUXI="' + stNatureauxi + '"' +
                 stCriteres + ' GROUP BY GL_TIERS, T_LIBELLE ' + stOrderBy;
   end;

end;

//Suppression des filles de la tob palmarès suivant le nombre de lignes
//à afficher par rupture
//Somme des Qtés et CA par rupture
procedure TOF_GraphPalm.CpteLigne(TobPal : TOB);
Var i_ind, iCpte : integer;
    TobPalAutre : TOB;
begin
if TRadioButton(GetControl('RMARGE')).Checked = True then
   if TRadioButton(GetControl('RMEILLEURES')).Checked = True then
         TobPal.Detail.sort('-GZA_MARGE')
   else  TobPal.Detail.sort('GZA_MARGE');
TobPalAutre := nil;
i_ind := 0;
InitMove(TobPal.Detail.count,'');
for iCpte := 0 to TobPal.Detail.Count -1 do
    begin
    MoveCur(False);
    if iCpte > StrToInt(GetControlText('SELECTION')) -1 then
        begin
        if TCheckBox(GetControl('RAUTRES')).Checked = True then
            begin
            if TobPalAutre = nil then
                begin
                TobPalAutre := Tob.Create ('GCTMPPAL', nil, -1);
                TobPalAutre.Dupliquer (TobPal.Detail[i_ind], true, true);
                TobPalAutre.PutValue ('GZA_PALMARES', 'Autres');
                TobPalAutre.PutValue ('GZA_LIBELLE', 'Autres');
                TobPalAutre.PutValue ('DIMENSIONS', '');
                end else
                begin
                TobPalAutre.PutValue ('GZA_QTEFACT',
                                      TobPalAutre.GetValue ('GZA_QTEFACT') +
                                            TobPal.Detail[i_ind].GetValue ('GZA_QTEFACT'));
                TobPalAutre.PutValue ('GZA_TOTALHT',
                                      TobPalAutre.GetValue ('GZA_TOTALHT') +
                                            TobPal.Detail[i_ind].GetValue ('GZA_TOTALHT'));
                TobPalAutre.PutValue ('GZA_MARGE',
                                      TobPalAutre.GetValue ('GZA_MARGE') +
                                            TobPal.Detail[i_ind].GetValue ('GZA_MARGE'));
                end;
            end;
        TobPal.Detail[i_ind].Free;
        end else inc(i_ind);
    if i_ind > TobPal.Detail.Count-1 then break;
    end;
if TobPalAutre <> nil then TobPalAutre.ChangeParent(TobPal,-1);
FiniMove;
end;

//Récupèration des critères et ajout au select
procedure TOF_GraphPalm.RecupCritere(Choix : integer);
var i_ind : integer;
    stFamille, stF, stcritFamille,condition : string;
begin

(*  gm
if GetControlText('NATUREPIECEG')='ZZ1' then
       begin   // mcd 10/10/02 ajout test ctxaffaire
       if ctxaffaire in V_PGI.PGIContexte then stCriteres := stCriteres + ' AND (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="AVC" OR GL_NATUREPIECEG="FRE")'
       else stCriteres := stCriteres + ' AND (GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="AVC" OR GL_NATUREPIECEG="AVS")';
       end
*)
Condition := ' AND ';
if not(ctxAffaire in V_PGI.PGIContexte)  then
    stcriteres := stcriteres + Condition + FactureAvoir_RecupSQL(GetControlText ('NATUREPIECEG'),'GL')
else
  Begin
  // gm 05/09/02  Prise en compte Facture reprise
  if GetControlText('NATUREPIECEG')='ZZ1' then
    stcriteres := stcriteres + Condition + '(GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="AVC" OR GL_NATUREPIECEG="AVS" OR GL_NATUREPIECEG="FRE")'
  else
    if GetControlText('NATUREPIECEG')='ZZ2' then
      stcriteres := stcriteres + Condition + '(GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="FRE")'
    else
      stcriteres := stcriteres + Condition + 'GL_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
  ENd;


if GetControlText('DATEPIECE') <> '' then
   stCriteres := stCriteres + ' AND GL_DATEPIECE>="'
                     + UsDateTime(StrToDate(GetControlText('DATEPIECE'))) + '"';

if GetControlText('DATEPIECE_') <> '' then
     stCriteres := stCriteres + ' AND GL_DATEPIECE<="'
                    + UsDateTime(StrToDate(GetControlText('DATEPIECE_'))) + '"';

for i_ind := 1 to 3 do
    begin
    stCritFamille := '';
    stFamille := GetControlText('FAMILLENIV' + IntToStr(i_ind));
    if stFamille = '<<Tous>>' then continue;
    stF := ReadTokenSt(stFamille);
    while stF <> '' do
        begin
        if stcritFamille = '' then stcritFamille := Concat('GL_FAMILLENIV',IntToStr(i_ind)) + '="' + stF + '"'
        else stcritFamille := stcritFamille + ' OR ' + Concat('GL_FAMILLENIV',IntToStr(i_ind)) + '="' + stF + '"';
        stF := ReadTokenSt(stFamille);
        end;
    if stcritFamille <> '' then stCriteres := stCriteres + ' AND (' + stcritFamille + ')';
    end;

if GetControlText('TYPEARTICLE') <> '' then
       stCriteres := stCriteres + ' AND GL_TYPEARTICLE="'
       + GetControlText('TYPEARTICLE') + '"';

if GetControlText('DEPOT') <> '' then
       stCriteres := stCriteres + ' AND GL_DEPOT="'
       + GetControlText('DEPOT') + '"';

if GetControlText('ETABLISSEMENT') <> '' then
       stCriteres := stCriteres + ' AND GL_ETABLISSEMENT="'
       + GetControlText('ETABLISSEMENT') + '"';
end;

procedure TOF_GraphPalm.RempliTOBPalmares(TobT : TOB; var TobPalmares : TOB);
var TobPal : TOB;
    PourcMarge : double;
    stDim, LibDim,GrilleDim,CodeDim : string;
    i_indDim : integer;
begin
MoveCur(False);
TobPal := TobPalmares.FindFirst(['GZA_PALMARES'], [TobT.GetValue('PALMARES')], False);
if TobPal = nil then
    begin
    TobPal := TOB.Create('GCTMPPAL',TobPalmares,-1);
    TobPal.PutValue('GZA_UTILISATEUR',V_PGI.USer);
    TobPal.PutValue('GZA_CLE',IntToStr(i_cle));
    TobPal.PutValue('GZA_LIBELLE',TOBT.GetValue('LIBELLE'));
    TobPal.PutValue('GZA_PALMARES',TOBT.GetValue('PALMARES'));
    TobPal.AddChampSup('DIMENSIONS',false);
    StDim:='';
    if TRadioButton(GetControl('RARTICLE')).Checked then
       if TobT.GetValue('GA_STATUTART')='DIM' then
          begin
          for i_indDim := 1 to MaxDimension do
              begin
              GrilleDim := TOBT.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
              CodeDim := TOBT.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
              LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
              if LibDim <> '' then
                 if StDim='' then StDim:=StDim + LibDim
                 else StDim := StDim + ' - ' + LibDim;
              end;
          end;
    TobPal.PutValue('DIMENSIONS',stDim);
    TobPal.PutValue('GZA_QTEFACT', 0);
    TobPal.PutValue('GZA_TOTALHT', 0);
    TobPal.PutValue('GZA_MARGE', 0);
    end;
TobPal.PutValue('GZA_QTEFACT', TobPal.GetValue('GZA_QTEFACT') + TOBT.GetValue('QTEFACT'));
TobPal.PutValue('GZA_TOTALHT', TobPal.GetValue('GZA_TOTALHT') + TOBT.GetValue('TOTALHT'));
if TRadioButton(GetControl('RTIERS')).Checked <> True then
    begin
    if TOBT.GetValue('TOTALHT')=0 then
        begin
        PourcMarge := 0;
        end else
        begin
        PourcMarge := ((TOBT.GetValue('TOTALHT')- TOBT.GetValue('TOTALHTACH'))
                / TOBT.GetValue('TOTALHT'))*100;
        end;
    TobPal.PutValue('GZA_MARGE', TobPal.GetValue('GZA_MARGE') + PourcMarge);
    end;
end;

procedure Tof_GraphPalm.ZoomGraphe (Sender: TCustomChart; Series: TChartSeries;
                                    ValueIndex: Integer; Button: TMouseButton;
                                    Shift: TShiftState; X, Y: Integer);
var F : TFGRS1;
    stArticle : string;
begin
F := TFGRS1(Ecran);
if Trim(F.FListe.Cells[0, ValueIndex + 1]) = 'Autres' then exit;
if TRadioButton(GetControl('RARTICLE')).Checked then
    begin
    stArticle := CodeArticleUnique2(F.FListe.Cells[0, ValueIndex + 1], '');
{$IFNDEF GPAO}
    AglLanceFiche ('GC', 'GCARTICLE', '', stArticle, 'ACTION=CONSULTATION');
{$ELSE}
	  V_PGI.DispatchTT(7, taConsult, stArticle, '', '');
{$ENDIF GPAO}
    end else
    begin
    if TRadioButton(GetControl('RTIERS')).Checked then
        begin
        AglLanceFiche ('GC', 'GCTIERS', '', F.Fliste.Cells[0, ValueIndex + 1],
                       'ACTION=CONSULTATION');
        end else
        begin
        AglLanceFiche ('GC', 'GCCOMMERCIAL', '', F.Fliste.Cells[0, ValueIndex + 1],
                       'ACTION=CONSULTATION');
        end;
    end;
end;

// ********************* register ******************************************
Procedure TOF_GraphPalm_CacheLibFam (parms:array of variant; nb: integer );
var F : TFORM;
    i_ind : integer;
begin
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCGRAPHPALM') then exit;
for i_ind := 1 to 3 do
    THMultiValComboBox(F.FindComponent('FAMILLENIV'+intToStr(i_ind))).Text := '';
TPanel(F.FindComponent('PFAMILLE')).Visible := False;
THValComboBox(F.FindComponent('TYPEARTICLE')).Visible := False;
THValComboBox(F.FindComponent('TYPEARTICLE')).Value := '';
THLabel(F.FindComponent('TTYPEARTICLE')).Visible := False;
end;

Procedure TOF_GraphPalm_ChangeAffichage (parms:array of variant; nb: integer );
var F : TFGRS1;
begin
F := TFGRS1 (Longint (Parms[0]));
if (F.Name <> 'GCGRAPHPALM') then exit;
if TCheckBox(F.FindComponent('RPOURCENTAGE')).Checked = True then
    begin
    F.FChart1.Legend.LegendStyle := lsValues;
    F.FChart1.Legend.TextStyle := ltsLeftPercent;
    end else
    begin
    F.FChart1.Legend.LegendStyle := lsValues;
    F.FChart1.Legend.TextStyle := ltsLeftValue;
    end;
end;

procedure InitTOFGraphPalm ();
begin
RegisterAglProc('GraphPalm_CacheLibFam', True , 0, TOF_GraphPalm_CacheLibFam);
RegisterAglProc('GraphPalm_ChangeAffichage', True , 0, TOF_GraphPalm_ChangeAffichage);
end;

Initialization
registerclasses([TOF_GraphPalm]);
InitTOFGraphPalm();

end.