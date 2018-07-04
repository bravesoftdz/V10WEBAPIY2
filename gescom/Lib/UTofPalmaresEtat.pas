unit UTofPalmaresEtat;

interface

uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls, graphics,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOB,UTOM,AglInit,Dialogs,
      M3FP, EntGC,Paramsoc,
{$IFDEF EAGLCLIENT}
      eQRS1,eFiche,MaineAGL,
{$ELSE}
       QRS1,HQuickRP, db,Fiche, dbTables,DBGrids,mul, Fe_Main,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}       
{$ENDIF}
      grids,UtilGc, HStatus,ExtCtrls, UTofEditStat ;


Type
     TOF_PalmaresEtat = Class (TOF)

        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad ; override ;

        function  CreateSQL(Choix : integer) : string ;
        procedure RecupCritere;
        procedure AjoutTable(NomTable: string);
        function CreerJointure(Rupture: string ; Num : integer) : integer;

        procedure RempliTOBPalmares(TobT : TOB ; var TobPalmares : TOB);
        function  RecupLibelle(NumRupt : integer ; Code : string) : string;
        procedure CpteLigne(TobPal : TOB);
        procedure CalculPourcentage(i_ind,i_ind2 : integer ; TotGen, QteGen : double ; TobPal : TOB);

     private

        TobFamille : TOB;
        Select   : string;
        Table    : string;
        Jointure : string;
        criteres : string;
        GroupBy  : string;
        OrderBy  : string;
        Condition : string;
        stNatureauxi : string;
        i_cle : integer;
     end;


implementation

procedure TOF_PalmaresEtat.OnArgument (stArgument : String ) ;
var Iind : integer;
begin
Inherited;
criteres := '';
THValComboBox(GetControl('TYPEARTICLE')).Value := 'MAR';
if stArgument = 'VEN' then
    begin
    stNatureauxi := 'CLI';
{$IFDEF NOMADE}
    THValComboBox(GetControl ('NATUREPIECEG')).Plus := 'AND GPP_VENTEACHAT="VEN"'
    + ' AND (GPP_NATUREPIECEG IN ("CC","DE") OR GPP_TRSFVENTE="X")' ;
{$ELSE}
    THValComboBox(GetControl ('NATUREPIECEG')).Plus := 'AND GPP_VENTEACHAT="VEN"';
{$ENDIF}
    TRadioButton(GetControl('RTIERS')).Caption := 'Client';
    TRadioButton(GetControl('RMARGE')).Enabled := True;
    TGroupBox(GetControl('GORDRE')).Caption := 'Ventes';
    //mcd 27/03/02*** Blocages des natures de pièces autorisées en affaires
    if (ctxAffaire in V_PGI.PGIContexte)  then
       begin
       SetCOntrolVisible('DEPOT',False);
       SetCOntrolVisible('TDEPOT',False);
       if (ctxScot in V_PGI.PGIContexte)  then
       	THValComboBox (Ecran.FindComponent('NATUREPIECEG')).Plus :=  'AND ((GPP_NATUREPIECEG="FAC") or (GPP_NATUREPIECEG="FPR")  or (GPP_NATUREPIECEG="APR") or (GPP_NATUREPIECEG="AVC") or (GPP_NATUREPIECEG="FRE"))';   //mcd 20/11/02 ajout fre
       end;
{$IFDEF NOMADE}
    SetControlText('NATUREPIECEG','CC') ; // Définition valeur par défaut
{$ELSE}
// Pourquoi ne pas faire le choix dans les filtres ? avec une autre fiche pour les achats ? JCF
    if (ctxMode in V_PGI.PGIContexte) then
    	SetControlText('NATUREPIECEG','FFO')
    else
    begin
       Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FAC');
       if Iind<0 then Iind:=0;

       if not(ctxAffaire in V_PGI.PGIContexte)  then
       begin
         FactureAvoir_AddItem(GetControl('NATUREPIECEG'),'VEN');
         setControlText('NATUREPIECEG','ZZV');
       end else  // gm 05/09/02 Prise en compte des Factures de reprise
       begin
 	 THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Avoir clients+Fac. Reprise');
         THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ1');
         SetControlText('NATUREPIECEG','ZZ1');
       end;
       if (ctxAffaire in V_PGI.PGIContexte)  then
       Begin  // gm 05/09/02 Prise en compte des Factures de reprise
         Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FRE');
         if Iind<0 then Iind:=0;
         THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Fac. Reprise');
         THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ2');
         SetControlText('NATUREPIECEG','ZZ2');
       end ;

    end ;
{$ENDIF}
    end
else
    begin
    stNatureauxi := 'FOU';
{$IFDEF NOMADE}
    THValComboBox(GetControl('NATUREPIECEG')).Plus := 'AND GPP_VENTEACHAT="ACH"'
    + ' AND (GPP_NATUREPIECEG="CF" OR GPP_TRSFACHAT="X")' ;
{$ELSE}
    THValComboBox(GetControl('NATUREPIECEG')).Plus := 'AND GPP_VENTEACHAT="ACH"';
{$ENDIF}
    TRadioButton(GetControl('RTIERS')).Caption := 'Fournisseur';
    TRadioButton(GetControl('RMARGE')).Enabled := True;
    TGroupBox(GetControl('GORDRE')).Caption := 'Achats';
{$IFDEF NOMADE}
    SetControlText('NATUREPIECEG','CF') ; // Définition valeur par défaut
{$ELSE}
    FactureAvoir_AddItem(GetControl('NATUREPIECEG'),'ACH');
    SetControlText('NATUREPIECEG','BLF') ;
{$ENDIF}
    SetControlProperty('GORDRE','Caption','Achats');
    SetControlProperty('RMEILLEURE','Caption','Les meilleurs');
    SetControlProperty('RMAUVAIS','Caption','Les plus petits');
    //*** Blocages des natures de pièces autorisées en affaires
    if (ctxAffaire in V_PGI.PGIContexte)  then
       begin
       SetCOntrolVisible('DEPOT',False);
       SetCOntrolVisible('TDEPOT',False);
       THValComboBox (Ecran.FindComponent('NATUREPIECEG')).Plus :='AND (GPP_VENTEACHAT="ACH") '+ AfPlusNAtureAchat;
       SetControlText ('NATUREPIECEG','FF');
       end;
    // **** Fin affaire ***
    end;
if (ctxMode in V_PGI.PGIContexte) then
   begin
        TRadioButton(GetControl('RARTICLE')).Visible := True;
        if (GetParamSoc('SO_GCMULTIDEPOTS')=False) then
          begin
	       THEdit(GetControl('CODE1')).Text:='AND CO_CODE<>"DEP"';
               THEdit(GetControl('CODE')).Text:='AND CO_CODE<>"DEP" AND CO_CODE<>"ZON"';
          end;
   end else TRadioButton(GetControl('RARTICLEDIM')).Caption := 'Article';
for iInd := 1 to 3 do
    begin
    SetControlCaption ('TFAMILLENIV' + IntToStr (iInd), RechDom('GCLIBFAMILLE','LF'+ IntToStr (iInd),false));
    end;

Inherited;
end;

procedure TOF_PalmaresEtat.OnLoad ;
var stOrdre : string;
begin
criteres := '';
if TRadioButton(GetControl('RARTICLE')).Checked then
//   begin
//   if (ctxMode in V_PGI.PGIContexte) then
//         SetControlText('XX_VARIABLE3','ARTICLEDIM')
//   else
     SetControlText('XX_VARIABLE3','ARTICLES');
//   end;
if TRadioButton(GetControl('RARTICLEDIM')).Checked then
//   begin
//   if (ctxMode in V_PGI.PGIContexte) then
         SetControlText('XX_VARIABLE3','ARTICLEDIM');
//   else  SetControlText('XX_VARIABLE3','ARTICLES');
//   end;

if TRadioButton(GetControl('RMEILLEURE')).Checked = True then
    begin
    if StrToInt(GetControlText('XX_VARIABLE4')) > 1 then
        begin
        if stNatureauxi = 'CLI' then
            stOrdre := 'Les ' + GetControlText('XX_VARIABLE4') + ' meilleures ventes'
        else stOrdre := 'Les ' + GetControlText('XX_VARIABLE4') + ' meilleurs achats';
        end else
        begin
        if stNatureauxi = 'CLI' then stOrdre := 'La meilleure vente'
        else stOrdre := 'Le meilleur achat';
        end;
    end else
    begin
    if StrToInt(GetControlText('XX_VARIABLE4')) > 1 then
        begin
        if stNatureauxi = 'CLI' then
            stOrdre := 'Les ' + GetControlText('XX_VARIABLE4') + ' plus petites ventes'
        else stOrdre := 'Les ' + GetControlText('XX_VARIABLE4') + ' plus petits achats';
        end else
        begin
        if stNatureauxi = 'CLI' then stOrdre := 'La plus petite vente'
        else stOrdre := 'Le plus petit achat';
        end;
    end;
if TRadioButton(GetControl('RCA')).Checked = true then
    begin
    SetControlText('XX_VARIABLE5', stOrdre + ' par chiffre d''affaire');
    end else
    if TRadioButton(GetControl('RQUANTITE')).Checked = true then
    begin
    SetControlText('XX_VARIABLE5', stOrdre + ' par quantité');
    end else
    begin
    SetControlText('XX_VARIABLE5', stOrdre + ' par marge');
    end;
Inherited;
end;


Procedure TOF_PalmaresEtat.OnUpdate ;
Var Requete : string;
    QSelect : TQuery;
    i_ind  :integer;
    TobTemp, TobPalmares : TOB;
begin
Inherited ;
ExecuteSQL('DELETE FROM GCTMPPAL WHERE GZA_UTILISATEUR = "'+V_PGI.USer+'"');
if GetControlText('NATUREPIECEG') = '' then
   begin
   PGIBox('Vous devez renseigner une nature de pièce','Palmarès');
   exit;
   end;
Select := '';
Table := '';
Jointure := '';
criteres := '';
GroupBy := '';
OrderBy := '';
Condition := ' AND ';
i_cle := 0;
if TRadioButton(GetControl('RARTICLE')).Checked = True then i_ind := 1
else if TRadioButton(GetControl('RARTICLEDIM')).Checked = True then i_ind := 2
     else if TRadioButton(GetControl('RREPRESENTANT')).Checked = True then i_ind := 3
          else i_ind := 4;

Requete := CreateSQL(i_ind);
QSelect := OpenSQL(Requete,True);
TobTemp := TOB.Create('',nil,-1);
TobPalmares := TOB.Create('GCTMPPAL',nil,-1);
QSelect.First;
while not QSelect.Eof do
    begin
    TobTemp.SelectDb ('', QSelect);
    RempliTOBPalmares(TobTemp, TobPalmares);
    QSelect.Next;
    inc(i_cle);
    end;
TobTemp.Free;
Ferme(QSelect);

if TobPalmares.Detail.Count > 0 then
   begin
   CpteLigne(TobPalmares);
   TobPalmares.InsertOrUpdateDB();
   end;
TobPalmares.free;
End;

procedure TOF_PalmaresEtat.OnClose ;
begin
Inherited;
ExecuteSQL('DELETE FROM GCTMPPAL WHERE GZA_UTILISATEUR = "'+V_PGI.USer+'"');
end;

////////////////////////////////////////////////////////////////////////////////
//********************  CONSTRUCTION DE LA REQUETE  ****************************
////////////////////////////////////////////////////////////////////////////////

function TOF_PalmaresEtat.CreateSQL(Choix : integer) : string ;
var StSelect, Rupture, TriRupt, Palmares : string;
    ChampsCalc, Ordre, Reste : string;
    Marge : string;
    i_ind, NumOrdre,cpt : integer;
begin
NumOrdre := 0;
//cpt :=0; inutile cf warning
TriRupt := '';
//Construction du Select et de l'éventuelle jointure si les champs ruptures
//sont renseignés
for i_ind:=1 to 2 do
    begin
    Rupture := GetControlText('RUPTURE' + IntToStr(i_ind));
    if Rupture <> '' then
       begin
       cpt:=CreerJointure(Rupture, i_ind);
       if TriRupt <> '' then TriRupt := TriRupt + ', ';
       TriRupt :=  TriRupt + Rupture ;
       Select := Select + ', ' + Rupture + ' AS RUPT' + IntToStr(i_ind) ;
       // Compteur du nombre de colonne rajouté dans le select pour une rupture
       inc(cpt);
       // cumul de cpt afin d'avoir le nombre de colonne total rajouté dans le select
       NumOrdre := NumOrdre+cpt;
       end;
    end;

if TriRupt <> '' then
   begin
   GroupBy := GroupBy + ', ' + TriRupt;
   TriRupt := TriRupt + ', ';
   end;

// Récupération des conditions du lanceur d'état et ajout du préfixe des champs,
//les tables n'étant pas les mêmes d'un Palmarès à l'autre
RecupCritere;
Ordre := ' ' + GetControlText('Ordre');

//Tri sur le CA ou sur la Quantité
if TRadioButton(GetControl('RQUANTITE')).Checked = True then
   OrderBy := ' ORDER BY ' + TriRupt + IntToStr(4 + NumOrdre) + Ordre
   else if TRadioButton(GetControl('RCA')).Checked = True then
   OrderBy := ' ORDER BY ' + TriRupt + IntToStr(3 + NumOrdre) + Ordre
   // DCA - FQ MODE 11029 - 5 colonnes maxi : order by 6 ne passe pas sous ORACLE
   //else OrderBy := ' ORDER BY ' + TriRupt + IntToStr(6 + NumOrdre) + Ordre;
   else OrderBy := ' ORDER BY ' + TriRupt + IntToStr(5 + NumOrdre) + Ordre;

Marge := 'GL_'+VH_GC.GCMargeArticle;
if (Marge = 'GL_PMA') or (Marge = 'GL_PMR') then Marge := Marge + 'P';
// BBI : Reste:='GL_QTEFACT-GL_QTERESTE';
Reste:='GL_QTERESTE';
ChampsCalc := ', SUM((GL_TOTALHT*('+Reste+'))/GL_QTEFACT) AS TOTALHT, '
              +'SUM('+Reste+') AS QTEFACT';
if choix <> 4 then
ChampsCalc := ChampsCalc + ',SUM('+ Marge +'*'+'('+Reste+')) AS TOTALHTACH';
Select := Select + ChampsCalc;

//On élimine les lignes de commentaires
// JS 06112003 criteres := criteres + ' AND GL_QTEFACT<>0 AND GL_TOTALHT<>0 AND GL_TYPELIGNE="ART" ';
// sinon ne prend pas en compte les lignes d'échantillons

criteres := criteres + ' AND GL_QTEFACT<>0 AND GL_TYPELIGNE="ART" ';

//Choix de la requête suivant s'il s'agit d'un Palmarès article,
//représentant ou client
case Choix  of
   1 : begin
       {if (ctxMode in V_PGI.PGIContexte) then
          begin }
          Palmares := 'CODEARTICLE';
          {end
       else Palmares := 'ARTICLE';  }
       StSelect:='SELECT GL_'+Palmares+' AS PALMARES, GA_LIBELLE AS LIBELLE'
       + Select +' FROM LIGNE , ARTICLE ' + Table + 'WHERE GL_ARTICLE=GA_ARTICLE'
       + Jointure + criteres + ' GROUP BY GL_'+Palmares+', GA_LIBELLE ' + GroupBy + OrderBy;
       end;
   2 : begin
       {if (ctxMode in V_PGI.PGIContexte) then
          begin
          Palmares := 'CODEARTICLE';
          end
       else }
       Palmares := 'ARTICLE';
       StSelect:='SELECT GL_'+Palmares+' AS PALMARES, GA_LIBELLE AS LIBELLE'
       + Select +' FROM LIGNE , ARTICLE ' + Table + 'WHERE GL_ARTICLE=GA_ARTICLE'
       + Jointure + criteres + ' GROUP BY GL_'+Palmares+', GA_LIBELLE ' + GroupBy + OrderBy;
       end;
   // DCA - FQ MODE 10320 - Inclusion ligne sans comercial
   {
   3 : StSelect:='SELECT GL_REPRESENTANT AS PALMARES, GCL_LIBELLE AS LIBELLE'
     + Select +' FROM LIGNE, COMMERCIAL ' + Table + 'WHERE GL_REPRESENTANT=GCL_COMMERCIAL'
     + Jointure + criteres + ' GROUP BY GL_REPRESENTANT, GCL_LIBELLE ' + GroupBy + OrderBy;
   }
   3 : StSelect:='SELECT GL_REPRESENTANT AS PALMARES, GCL_LIBELLE AS LIBELLE'
     + Select +' FROM LIGNE ' + Table + 'LEFT JOIN COMMERCIAL ON GCL_COMMERCIAL=GL_REPRESENTANT WHERE GL_NATUREPIECEG=GL_NATUREPIECEG'
     + Jointure + criteres + ' GROUP BY GL_REPRESENTANT, GCL_LIBELLE ' + GroupBy + OrderBy;

   4 : StSelect:='SELECT GL_TIERS AS PALMARES, T_LIBELLE AS LIBELLE'
     + Select +' FROM LIGNE, TIERS ' + Table + 'WHERE GL_TIERS=T_TIERS AND ' +
     'T_NATUREAUXI="' + stNatureauxi + '"' +
     Jointure + criteres + ' GROUP BY GL_TIERS, T_LIBELLE ' + GroupBy + OrderBy;
   end;

Result := StSelect;
end;

//Récupèration des critères et ajout au select
procedure TOF_PalmaresEtat.RecupCritere;
var i_ind : integer;
    stFamille, stF, stcritFamille : string;
begin

if not(ctxAffaire in V_PGI.PGIContexte)  then
    criteres := criteres + Condition + FactureAvoir_RecupSQL(GetControlText('NATUREPIECEG'),'GL')
else
  Begin
  // gm 05/09/02  Prise en compte Facture reprise
  if GetControlText('NATUREPIECEG')='ZZ1' then
    criteres := criteres + Condition + '(GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="AVC" OR GL_NATUREPIECEG="AVS" OR GL_NATUREPIECEG="FRE")'
  else
    if GetControlText('NATUREPIECEG')='ZZ2' then
      criteres := criteres + Condition + '(GL_NATUREPIECEG="FAC" OR GL_NATUREPIECEG="FRE")'
    else
      criteres := criteres + Condition + 'GL_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
  ENd;



if GetControlText('DATEPIECE') <> '' then
   criteres := criteres + Condition + 'GL_DATEPIECE>="'
                     + UsDateTime(StrToDate(GetControlText('DATEPIECE'))) + '"';

if GetControlText('DATEPIECE_') <> '' then
     criteres := criteres + Condition + 'GL_DATEPIECE<="'
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
    if stcritFamille <> '' then criteres := criteres + Condition + '(' + stcritFamille + ')';
    end;

if GetControlText('TYPEARTICLE') <> '' then
       criteres := criteres + Condition + 'GL_TYPEARTICLE="'
       + GetControlText('TYPEARTICLE') + '"';

if GetControlText('DEPOT') <> '' then
       criteres := criteres + Condition + 'GL_DEPOT="'
       + GetControlText('DEPOT') + '"';

if GetControlText('ETABLISSEMENT') <> '' then
       criteres := criteres + Condition + 'GL_ETABLISSEMENT="'
       + GetControlText('ETABLISSEMENT') + '"';

end;


//Création des jointures et ajout des tables
// Cette fonction renvoie le nombre de colonne ajouté dans le select
function TOF_PalmaresEtat.CreerJointure(Rupture: string ; Num : integer) : integer;
Var  QSelect : TQuery;
begin
result:=-1;
if pos('FAMILLENIV',Rupture)<>0 then
   begin
      if TobFamille = nil then
        begin
             QSelect := OpenSQL('SELECT CC_TYPE, CC_CODE, CC_LIBELLE FROM CHOIXCOD WHERE'
                       + ' CC_TYPE LIKE "FN%"',false);
             TobFamille := TOB.Create('CHOIXCOD',nil,-1);
             if not QSelect.EOF then
                TobFamille.LoadDetailDB('','','',QSelect,false);
             Ferme(QSelect);
        end;
      result:=0;
   end;

if pos('ZONECOM',Rupture)<>0 then
   begin
        if TRadioButton(GetControl('RREPRESENTANT')).Checked = false then
          begin
               AjoutTable('COMMERCIAL');
               Jointure := Jointure + Condition + 'GL_REPRESENTANT=GCL_COMMERCIAL ';
          end;

          AjoutTable('CHOIXCOD');
          Jointure := Jointure + ' AND ' + RUPTURE + '=CC_CODE ';
          criteres := criteres + ' AND (CC_TYPE="GZC" OR CC_TYPE="") ' ;
          Select := Select + ', CC_LIBELLE AS LIB' + Concat('RUPT',IntToStr(num));
          GroupBy:= GroupBy  + ', CC_LIBELLE';
          result:=1;
          exit;
   end;

if pos('ETABLISSEMENT',Rupture)<>0 then
   begin
        AjoutTable('ETABLISS');
        Jointure := Jointure + Condition + 'GL_ETABLISSEMENT=ET_ETABLISSEMENT ';
        Select := Select + ', ET_LIBELLE AS LIB' + Concat('RUPT',IntToStr(num));
        GroupBy:= GroupBy  + ', ET_LIBELLE';
        result:=1;
        exit;
   end;

if pos('DEPOT',Rupture)<>0 then
   begin
   AjoutTable('DEPOTS');
   Jointure := Jointure + Condition + 'GL_DEPOT=GDE_DEPOT ';
   Select := Select + ', GDE_LIBELLE AS LIB' + Concat('RUPT',IntToStr(num));
   GroupBy:= GroupBy  + ', GDE_LIBELLE';
   result:=1;
   exit;
   end;
end;


procedure TOF_PalmaresEtat.AjoutTable(NomTable: string);
begin
if Pos(NomTable,Table) = 0 then
      Table := Table + ', ' + NomTable + ' ';
end;

////////////////////////////////////////////////////////////////////////////////
//***  GESTION DE LA TOB PALMARES AVANT INSERTION DANS LA TABLE TEMPORAIRE  ****
////////////////////////////////////////////////////////////////////////////////

procedure TOF_PalmaresEtat.RempliTOBPalmares(TobT : TOB ; var TobPalmares : TOB);
var
    TobPal : TOB;
    i_ind2 : integer;
    PourcMarge : double;
    FindCle : string;
    FindValeur : variant;
begin
FindCle:= 'GZA_PALMARES';
FindValeur := TobT.GetValue('PALMARES');
for i_ind2 := 1 to 2 do
    if  GetControlText('XX_RUPTURE' + IntToStr(i_ind2)) <> '' then
       begin
       FindCle:= FindCle + ',GZA_RUPT'+ IntToStr(i_ind2);
       FindValeur := FindValeur + ',' +TOBT.GetValue('RUPT'+ IntToStr(i_ind2));
       end;
TobPal := TobPalmares.FindFirst([FindCle],[FindValeur], False);
MoveCur(False);
if TobPal = nil then
   begin
   TobPal := TOB.Create('GCTMPPAL',TobPalmares,-1);
   TobPal.InitValeurs;
   TobPal.PutValue('GZA_UTILISATEUR',V_PGI.USer);
   TobPal.PutValue('GZA_CLE',IntToStr(i_cle));
   TobPal.PutValue('GZA_LIBELLE',TOBT.GetValue('LIBELLE'));
   TobPal.PutValue('GZA_PALMARES',TOBT.GetValue('PALMARES'));
   for i_ind2 := 1 to 2 do
      begin
      if  GetControlText('XX_RUPTURE' + IntToStr(i_ind2)) <> '' then
         begin
         TobPal.PutValue('GZA_RUPT' + IntToStr(i_ind2),TOBT.GetValue('RUPT'+ IntToStr(i_ind2)));
         if pos('FAMILLENIV',GetControlText('RUPTURE'+ IntToStr(i_ind2))) <> 0 then
         TobPal.PutValue('GZA_LIBRUPT' + IntToStr(i_ind2),
              RecupLibelle(i_ind2,TobPal.GetValue('GZA_RUPT'+ IntToStr(i_ind2))))
         else TobPal.PutValue('GZA_LIBRUPT' + IntToStr(i_ind2),
                        TOBT.GetValue('LIBRUPT'+ IntToStr(i_ind2)));
         end;
      end;
   end;
TobPal.PutValue('GZA_QTEFACT',TobPal.GetValue('GZA_QTEFACT')+TOBT.GetValue('QTEFACT'));
TobPal.PutValue('GZA_TOTALHT',TobPal.GetValue('GZA_TOTALHT')+TOBT.GetValue('TOTALHT'));

if TRadioButton(GetControl('RTIERS')).Checked = false then
   begin
   if TOBT.GetValue('TOTALHT')=0 then
      begin
      PourcMarge := 0;
      end else
      begin
      PourcMarge := ((TOBT.GetValue('TOTALHT')- TOBT.GetValue('TOTALHTACH'))
              / TOBT.GetValue('TOTALHT'))*100;
      end;
   TobPal.PutValue('GZA_MARGE',TobPal.GetValue('GZA_MARGE')+PourcMarge);
   end;
end;

//Recherche des libellés des ruptures familles d'articles
function TOF_PalmaresEtat.RecupLibelle(NumRupt : integer ; Code : string) : string;
Var i_ind : integer;
begin
i_ind := StrToInt(Copy(GetControlText('RUPTURE'+IntToStr(NumRupt)),14,1));
TobFamille.PutValue('CC_TYPE',Concat('FN',IntToStr(i_ind)));
TobFamille.PutValue('CC_CODE',Code);
TobFamille.ChargeCle1;
if not TobFamille.LoadDB then
   Result := ''
   else Result := TobFamille.GetValue('CC_LIBELLE');
end;

//Suppression des filles de la tob palmarès suivant le nombre de lignes
//à afficher par rupture
//Somme des Qtés et CA par rupture
procedure TOF_PalmaresEtat.CpteLigne(TobPal : TOB);
Var ORupt1, ORupt2 : string;
    i_ind, i_ind2, Cpte :integer;
    TotGen, QteGen : double;
begin
if TRadioButton(GetControl('RMARGE')).Checked = True then
   if TRadioButton(GetControl('RMEILLEURE')).Checked = True then
      TobPal.Detail.sort('GZA_RUPT1;GZA_RUPT2;-GZA_MARGE')
   else  TobPal.Detail.sort('GZA_RUPT1;GZA_RUPT2;GZA_MARGE');
i_ind := 0;
InitMove(TobPal.Detail.count,'');
repeat
  Cpte := 0;
  TotGen := 0;
  QteGen := 0;
  i_ind2:=0;
  ORupt1:=TobPal.Detail[i_ind].GetValue('GZA_RUPT1');
  ORupt2:=TobPal.Detail[i_ind].GetValue('GZA_RUPT2');
  while     (ORupt1 = TobPal.Detail[i_ind].GetValue('GZA_RUPT1'))
      and   (ORupt2 = TobPal.Detail[i_ind].GetValue('GZA_RUPT2')) do
    begin
    MoveCur(False);
    inc(Cpte);
    TotGen := TotGen + StrToFloat(TobPal.Detail[i_ind].GetValue('GZA_TOTALHT'));
    QteGen := QteGen + StrToFloat(TobPal.Detail[i_ind].GetValue('GZA_QTEFACT'));
    if Cpte > StrToInt(GetControlText('XX_VARIABLE4')) then
           TobPal.Detail[i_ind].free
    else
        begin
        inc(i_ind);
        inc(i_ind2);
        end;
    if i_ind > TobPal.Detail.Count-1 then break;
    end;
    CalculPourcentage(i_ind-1,i_ind2,TotGen, QteGen, TobPal);
    TobPal.Detail[i_ind-1].PutValue('GZA_TOTGEN',TotGen);
    TobPal.Detail[i_ind-1].PutValue('GZA_QTEGEN',QteGen);
until i_ind > TobPal.Detail.Count-1;
FiniMove;
end;

//Calcul du pourcentage Qté et CA
procedure TOF_PalmaresEtat.CalculPourcentage(i_ind,i_ind2 : integer ; TotGen, QteGen : double ; TobPal : TOB);
var PQte, PTot : double;
    i_indice : integer;
begin
for i_indice:= i_ind downto (i_ind+1-i_ind2) do
  begin
    if QteGen <> 0 then
    begin
      PQte := (TobPal.Detail[i_indice].GetValue('GZA_QTEFACT')/QteGen)*100;
      TobPal.Detail[i_indice].PutValue('GZA_POURCENTQTE',PQte);
    end;
    if TotGen <> 0 then
    begin
      PTot := (TobPal.Detail[i_indice].GetValue('GZA_TOTALHT')/TotGen)*100;
      TobPal.Detail[i_indice].PutValue('GZA_POURCENTTOT',PTot);
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
//****************  CONTROLE DES RUPTURES PARAMETRABLES  ***********************
////////////////////////////////////////////////////////////////////////////////

Procedure TOF_PalmaresEtat_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCPALMARES') then exit;
Indice := Integer (Parms[1]);
St_Plus := THEdit(F.FindComponent('CODE')).Text;
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text);
For i_ind := 1 to 2 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value);
    If St <> '' then
       begin
       St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"' ;
       end;
    END;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Plus := St_Plus;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value := St_Value;
THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Text := St_Text ;
END;

Procedure TOF_PalmaresEtat_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
    Rupt : THValComboBox;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCPALMARES') then exit;
Indice := Integer (Parms[1]);
St_Plus := THEdit(F.FindComponent('CODE')).Text;
St_Value := string (THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice))).Value);
Rupt:= THValComboBox (F.FindComponent('RUPT'+InttoStr(indice)));
if St_Value = '' then
    BEGIN
    For i_ind := Indice to 2 do
        begin
        TCheckBox(F.FindComponent('SAUTPAGE'+InttoStr(i_ind))).Checked:=False;
        TCheckBox(F.FindComponent('SAUTPAGE'+InttoStr(i_ind))).Enabled:=False;
        end;
    THEdit (F.FindComponent('RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(Indice))).Text := '';
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := '';
    Rupt.Plus := St_Plus;
    Rupt.Text := '';
    For i_ind := Indice + 1 to 2 do
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Plus := '';
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Enabled := False;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(i_ind))).Color := clBtnFace;
        THEdit (F.FindComponent('RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_RUPTURE'+InttoStr(i_ind))).Text := '';
        THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
        END;
    END else
    BEGIN
    if Indice < 2 then
        BEGIN
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('RUPT'+InttoStr(Indice + 1))).Color := clWindow;
        END;
    TCheckBox(F.FindComponent('SAUTPAGE'+InttoStr(Indice))).Enabled:=True;
    THEdit (F.FindComponent('RUPTURE'+InttoStr(Indice))).Text :=
            RechDom ('GCGROUPPALM', St_Value, True);
    if pos('LF',Rupt.Value) <>0 then
            THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text :=
            RechDom('GCLIBFAMILLE',Rupt.Value,False)
    else    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(Indice))).Text := Rupt.Text;
    END;
END;

////////////////////////////////////////////////////////////////////////////////
Procedure TOF_PalmaresEtat_CacheLibFam (parms:array of variant; nb: integer ) ;
var F : TFORM;
    i_ind : integer;
begin
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCPALMARES') then exit;
for i_ind := 1 to 3 do
    THMultiValComboBox(F.FindComponent('FAMILLENIV'+intToStr(i_ind))).Text := '';
TPanel(F.FindComponent('PFAMILLE')).Visible := False;
THValComboBox(F.FindComponent('TYPEARTICLE')).Visible := False;
THValComboBox(F.FindComponent('TYPEARTICLE')).Value := '';
THLabel(F.FindComponent('TTYPEARTICLE')).Visible := False;
for i_ind := 1 to 2 do
    THLabel(F.FindComponent('LIBFAMILLE'+IntTostr(i_ind))).Caption:= '';
for i_ind := 7 to 9 do
    THEdit (F.FindComponent('XX_VARIABLE'+InttoStr(i_ind))).Text := '';
if TRadioButton(F.FindComponent('RTIERS')).Checked = True then
    begin
    if TOF_PalmaresEtat(TFQRS1(F).LaTOF).stNatureauxi = 'CLI' then
        THEdit(F.FindComponent('XX_VARIABLE3')).Text := 'CLIENTS'
    else THEdit(F.FindComponent('XX_VARIABLE3')).Text := 'FOURNISSEURS';
    end;
end;

/////////////////////////////////////////////////////////////////////////////

procedure InitTOFPalmaresEtat ();
begin
RegisterAglProc('PalmaresEtat_ChangeGroup', True , 1, TOF_PalmaresEtat_ChangeGroup);
RegisterAglProc('PalmaresEtat_AffectGroup', True , 1, TOF_PalmaresEtat_AffectGroup);
RegisterAglProc('PalmaresEtat_CacheLibFam', True , 0, TOF_PalmaresEtat_CacheLibFam);
end;

Initialization
registerclasses([TOF_PalmaresEtat]);
InitTOFPalmaresEtat();

end.
