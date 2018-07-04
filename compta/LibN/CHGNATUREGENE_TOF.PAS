{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/10/2001
Modifié le ... : 16/11/2006
Description .. : Source TOF de la TABLE : CHGNATUREGENE ()
Suite ........ : --> FQ 19113 - CA - 16/11/2006 - Pour ne jamais avoir la
Suite ........ : nature à ""
Mots clefs ... : TOF;CHGNATUREGENE
*****************************************************************}
Unit CHGNATUREGENE_TOF ;

Interface

Uses Controls, Classes, sysutils,
{$IFDEF EAGLCLIENT}
     Maineagl,
     eMul,
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     HDB,
     Mul,
     FE_Main,
{$ENDIF}
     ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF, HTB97,
     Paramsoc, HQry, UTob, CPGeneraux_TOM,
     lookup,
     UTobdebug,
     Ent1;

procedure CPLanceFiche_ChgNatureGene;

Type
  TOF_CHGNATUREGENE = Class (TOF)
    NatureOri, NewNature : THValComboBox ;
    TobGene : Tob ;
    Fiche : TFMul ;
{$IFDEF EAGLCLIENT}
    ListeCpt : THGrid;
{$ELSE}
    ListeCpt : THDBGrid;
{$ENDIF}
    QListeCpt : TQuery;
    Pages : TpageControl ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure NatureOriOnChange(sender: TObject);
    procedure NewNatureOnChange(sender: TObject);
    procedure BLancerOnClick(sender: TObject); {JP 26/08/04, FQ 14058}
    procedure FListeOnDblClick(sender: TObject);
  private
    procedure RecupMultiSelect;
    procedure TraitementCompte;
    procedure PrechargeOrigines;
    procedure MarqueOrigine;
    procedure VireInutiles;
    procedure RechTiers(CpteAuxi: string ;var Lettrable : boolean ; var ModeRegle : string);
    function RechModePaie(ModeRegle: string): string;
    procedure ModifCompte(TobC: Tob; ValLettrage, SuiviTreso, TVADefaut : string );
    function  ModifEcriture(CptGene, CptAuxi, ValLettrage, ValPointage, ModePaie : string; LetAuxi : Boolean) : Integer;
    function RechTVA: string;
    function ExisteMouvementsAnterieurs(TobDet: Tob): Boolean;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  CPTypeCons,
  {$ENDIF MODENT1}
  ulibecriture, SoldeCpt;

var
  Titre : string;

procedure CPLanceFiche_ChgNatureGene;
begin
  AGLLanceFiche('CP' , 'CHGNATUREGENE', '', '' , '' );
end;


procedure TOF_CHGNATUREGENE.NatureOriOnChange ( sender : TObject) ;
begin
  // FQ 19113 - CA - 16/11/2006 - Pour ne jamais avoir la nature à ""
  if (NatureOri.Value='') then NatureOri.Value := 'DIV';
    // mise à jour des tablettes en fonction de la nature d'origine choisie
    SetControlText('XX_WHERE','') ;
{ BPY 29/01/2003 Correction de la fiche 10541 }
//    if NatureOri.Value = 'DIV' then   NewNature.Plus := 'AND (CO_CODE<>"IMO") AND (CO_CODE<>"EXT") AND (CO_CODE<>"DIV")' ;
    if NatureOri.Value = 'DIV' then   NewNature.Plus := 'AND (CO_CODE<>"EXT") AND (CO_CODE<>"DIV")' ;
{ Fin BPY }
    if Copy(NatureOri.Value,1,2)='TI' then
    begin
        NewNature.Plus := 'AND (CO_CODE LIKE "CO%" OR CO_CODE="EXT"  OR CO_CODE="DIV" OR CO_CODE="BQE" OR CO_CODE="CAI") ' ;
//        SetControlText('XX_WHERE','G_LETTRABLE="X"') ; // BPY le 12/08/2003 puis 09/10/2003 : bug n°12349 puis 12489
    end ;
    NewNature.Value := '' ;
end ;

procedure TOF_CHGNATUREGENE.NewNatureOnChange ( sender : TObject ) ;
begin
  // saisie ou non du compte auxiliaire suivant la nature de compte choisie
  SetControlText('NEWAUXI','') ;
  SetControlEnabled('NEWAUXI',False) ;
  SetControlEnabled('TAUXI',False) ;
//gv 11/06/02
{  if (NatureOri.Value = 'DIV') and (Copy(NewNature.Value,1,2)='CO') then
    begin
    SetControlEnabled('NEWAUXI',true) ;
    SetControlEnabled('TAUXI',true) ;
    SetControlProperty('NEWAUXI','DATATYPE','TZTDIVERS') ;
    end ;}
//gv 11/06/02
  if (not (Copy(NatureOri.Value,1,2)='CO')) and (Copy(NewNature.Value,1,2)='CO') then
    begin
    SetControlEnabled('NEWAUXI',true) ;
    SetControlEnabled('TAUXI',true) ;
    if NewNature.Value = 'COC' then SetControlProperty('NEWAUXI','DATATYPE','TZTCLIENT') ;
    if NewNature.Value = 'COF' then SetControlProperty('NEWAUXI','DATATYPE','TZTFOURN') ;
    if NewNature.Value = 'COS' then SetControlProperty('NEWAUXI','DATATYPE','TZTSALARIE') ;
    if NewNature.Value = 'COD' then SetControlProperty('NEWAUXI','DATATYPE','TZTDIVERS') ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /    
Description .. : Recherche des comptes suivant les critères du MUL.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CHGNATUREGENE.PrechargeOrigines ;
Var Q : TQuery ;
    ReqSql, St : String ;
begin
  // recherche des comptes
  ReqSql := 'SELECT G_GENERAL,G_NATUREGENE,G_COLLECTIF,G_LETTRABLE, G_POINTABLE, G_REGIMETVA, G_SUIVITRESO, G_DERNLETTRAGE, ' ; // gv 11/06/02 Rajout G_SUIVITRESO
  ReqSql := ReqSql + ' G_MODEREGLE, G_TVAENCAISSEMENT, G_TVA FROM GENERAUX ' ;                                  // SG6 11.03.05 rajout de G_DERNLETTRAGE FQ 12309
  St:=AnsiUpperCase(RecupWhereCritere(Pages)) ;
  ReqSql := ReqSql + St + ' ORDER BY G_GENERAL' ;
  Q := OpenSQL(ReqSql, True) ;
  TobGene.LoadDetailDB('GENERAUX','','',Q,False,True) ;
  Ferme(Q) ;
  if TOBGene.Detail.Count>0 then TOBGene.Detail[0].AddChampSup('MARQUE',True) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /    
Description .. : Marquage dans la tob des comptes sélectionnés.
Mots clefs ... :
*****************************************************************}
procedure TOF_CHGNATUREGENE.MarqueOrigine ;
var TobC : Tob ;
begin
  TOBC := TobGene.FindFirst(['G_GENERAL'],[QListeCpt.FindField('G_GENERAL').AsString],False) ;
  if TOBC<>Nil then TOBC.PutValue('MARQUE','X') ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /    
Description .. : Suppression de la tob des comptes non sélectionnés.
Mots clefs ... :
*****************************************************************}
procedure TOF_CHGNATUREGENE.VireInutiles ;
var idxCpt : integer ;
    TobC : Tob ;
begin
  for idxCpt := TobGene.Detail.Count -1 downto 0 do
    begin
    TobC := TobGene.Detail[idxCpt] ;
    if TobC.GetValue('MARQUE') <> 'X' then begin TobC.Free ; end ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /    
Description .. : Construction de la tob des comptes d'après la sélection.
Mots clefs ... :
*****************************************************************}
procedure TOF_CHGNATUREGENE.RecupMultiSelect ;
var idxCpt : integer ;
begin
  // récupération des comptes sélectionnés de l'écran
  TobGene.ClearDetail ;
  PrechargeOrigines ;
  if Not ListeCpt.AllSelected then
    begin
    for idxCpt:=0 to ListeCpt.NbSelected-1 do
      begin
      ListeCpt.GotoLeBookmark(idxCpt) ;
      {$IFDEF EAGLCLIENT}
      Fiche.Q.TQ.Seek(Fiche.FListe.Row-1);
      {$ENDIF}
      MarqueOrigine ;
      end ;
    VireInutiles ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /    
Description .. : Recherche des différents éléments du compte auxiliaire 
Suite ........ : saisi.
Mots clefs ... :
*****************************************************************}
procedure TOF_CHGNATUREGENE.RechTiers ( CpteAuxi : string ; var Lettrable : boolean ; var ModeRegle : string );
var Q : TQuery ;
    Req : String ;
begin
  {JP 30/06/05 : FQ 15364 : violation d'accès en CWAS si le Query est vide}
  if CpteAuxi = '' then Exit;
  // récupère les données sur le lettrage et mode de règlement du tiers
  Req := 'SELECT T_LETTRABLE, T_MODEREGLE, T_TOTDEBP, T_TOTCREP, T_TOTDEBE, T_TOTCREE, ' ;
  Req := Req +'T_TOTDEBS, T_TOTCRES FROM TIERS WHERE T_AUXILIAIRE="'+CpteAuxi+'"' ;
  Q := OpenSql (Req,true ) ;
  {JP 30/06/05 : FQ 15364 : violation d'accès en CWAS si le Query est vide}
  if not Q.EOF then begin
    Lettrable := Q.Fields[0].asString = 'X' ;
    ModeRegle := Q.Fields[1].asString ;
  end;
  Ferme(Q) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /    
Description .. : recherche le mode de paiement correspondant au mode de 
Suite ........ : règlement
Mots clefs ... : 
*****************************************************************}
function TOF_CHGNATUREGENE.RechModePaie ( ModeRegle : string ) : string ;
var Q : TQuery ;
begin
  // recherche le mode de paiement correspondant au mode de règlement
  result := '' ;
  Q := OpenSql ('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'+ModeRegle+'"',true ) ;
  result := Q.Fields[0].asString ;
  Ferme(Q) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /
Description .. : Modification des zones du compte suivant la nature choisie.
Mots clefs ... : 
*****************************************************************}
procedure TOF_CHGNATUREGENE.ModifCompte ( TobC : Tob ; ValLettrage, SuiviTreso, TVADefaut : string ) ;
begin
  // mise à jour des zones du compte suivant la nature
  TobC.PutValue('G_NATUREGENE',NewNature.Value) ;
  If ValLettrage<>'' then TobC.PutValue('G_LETTRABLE',ValLettrage)
  //SG6 11.03.05 FQ 12309
  else TobC.PutValue('G_DERNLETTRAGE','');

  If Copy(NewNature.Value,1,2)='CO' Then TobC.PutValue('G_LETTRABLE', '-') ;

  // les comptes Tiers, collectifs, charge et produit ne sont pas pointables
  If (Copy(NewNature.Value,1,2)='TI') Or (Copy(NewNature.Value,1,2)='CO') or
     (NewNature.Value='CHA') or ( NewNature.Value='PRO') then TobC.PutValue('G_POINTABLE','-') ;

  // mise à jour de TVA par defaut si CHArge ou PROduit
  if (NewNature.Value='CHA') or ( NewNature.Value='PRO') then TobC.PutValue('G_TVA',TVADefaut) ;

  // comptes collectifs
  if Copy(NewNature.Value,1,2) = 'CO' then TobC.PutValue('G_COLLECTIF', 'X' ) ;

  // mise à jour des données TVA et règlement pour les Tiers
  if Copy(NewNature.Value,1,2) = 'TI' then
    begin
    TobC.PutValue('G_REGIMETVA', VH^.RegimeDefaut ) ;
    TobC.PutValue('G_MODEREGLE', GetParamsocSecur('SO_GCMODEREGLEDEFAUT','')) ;
    TobC.PutValue('G_TVAENCAISSEMENT', GetParamsocSecur('SO_TVAENCAISSEMENT','') ) ;
    end ;
  if Copy(NatureOri.Value,1,2) = 'TI' then
    begin
    TobC.PutValue('G_REGIMETVA', '' ) ;
    TobC.PutValue('G_MODEREGLE', '' ) ;
    TobC.PutValue('G_TVAENCAISSEMENT', '' ) ;
    end ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /
Description .. : Modification des écritures comptables en fonction de la
Suite ........ : nouvelle nature choisie.
Mots clefs ... :
*****************************************************************}
function TOF_CHGNATUREGENE.ModifEcriture(CptGene, CptAuxi, ValLettrage, ValPointage, ModePaie : string; LetAuxi : Boolean) : Integer;
var
  SQL : string;
begin
  SQL := 'UPDATE ECRITURE SET E_PAQUETREVISION = 1';
  if CptAuxi <> '' then SQL := SQL + ', E_AUXILIAIRE = "' + CptAuxi + '"';
  if ValLettrage = 'X' then
    SQL := SQL + ', E_MODEPAIE = IIF(((E_MODEPAIE = "") OR (E_MODEPAIE IS NULL)), "' + ModePaie + '", E_MODEPAIE)';
  if ValPointage = '-' then
    SQL := SQL + ', E_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '", E_REFPOINTAGE = "", E_NATURETRESO = ""';

  {JP 19/10/05 : FQ 16797 : Si la nouvelle nature est Divers et que l'ancien compte était lettrable
                 on récupère les informations de lettrage partiellement
   REMARQUE : j'ai l'impression que cette méthode n'est pas complétement abouti :
              quid de ValPointage}
  if (NewNature.Value = 'DIV') and (ValLettrage = 'X') then begin
    { FQ 19086 BVE 12.04.07 : On veux conserver la date d'echeance }
    SQL := SQL + ', E_DATEECHEANCE = E_DATECOMPTABLE, E_ORIGINEPAIEMENT = "' + UsDateTime(iDate1900) + '"';
    SQL := SQL + ', E_DATERELANCE = "' + UsDateTime(iDate1900) + '", E_ENCAISSEMENT = "RIE"';
    {Pour la Tréso, on ne fait rien, car cela concerne les flux prévisionnels et à priori on ne touche pas à
     la clef et au lettrage
     SQL := SQL + ', E_TRESOSYNCHRO = "RIE"';}
    SQL := SQL + ', E_TYPEMVT = IIF(E_TYPEMVT = "TTC", "DIV", E_TYPEMVT)';
  end

  // mise à jour du lettrage si besoin
  else if (ValLettrage = 'X') or ((CptAuxi <> '') and LetAuxi) then begin
    SQL := SQL + ', E_LETTRAGE = "", E_LETTRAGEDEV = "-", E_ECHE = "X", E_ETATLETTRAGE = "AL"';
    SQL := SQL + ', E_COUVERTURE = 0, E_COUVERTUREDEV = 0, E_DATEECHEANCE = E_DATECOMPTABLE';
    SQL := SQL + ', E_DATEVALEUR = E_DATECOMPTABLE, E_DATEPAQUETMAX = E_DATECOMPTABLE, E_DATEPAQUETMIN = E_DATECOMPTABLE';
    SQL := SQL + ', E_NUMECHE = IIF(E_NUMECHE = 0, 1, E_NUMECHE)';
  end

  else if (ValLettrage = '-') or ((CptAuxi <> '') and not LetAuxi) then begin
    SQL := SQL + ', E_LETTRAGE = "", E_LETTRAGEDEV = "-", E_ECHE = "-", E_ETATLETTRAGE = "RI"';
    SQL := SQL + ', E_COUVERTURE = 0, E_COUVERTUREDEV = 0, E_DATEPAQUETMAX = "' + UsDateTime(iDate1900) + '"';
    SQL := SQL + ', E_DATEPAQUETMIN = "' + UsDateTime(iDate1900) + '", E_DATEECHEANCE = "' + UsDateTime(iDate1900) + '"';
    SQL := SQL + ', E_NUMECHE = 0, E_MODEPAIE = "", E_ENCAISSEMENT = "RIE"';
  end;

  SQL := SQL + ' WHERE E_GENERAL = "' + CptGene + '"';
  Result := ExecuteSQL(SQL);
  {mise à jour du compte auxiliaire sur analytique si besoin}
  if (Result > 0) and (CptAuxi <> '') then
    ExecuteSQL('UPDATE ANALYTIQ SET Y_AUXILIAIRE = "' + CptAuxi + '" WHERE Y_GENERAL = "' + CptGene + '"');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /
Description .. : Recherche du premier code TVA
Mots clefs ... :
*****************************************************************}
function TOF_CHGNATUREGENE.RechTVA : string ;
var Req : string ;
    Q : TQuery ;
begin
  result := '' ;
  Req := 'SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="TX1"' ;
  Q := OpenSql(Req,true) ;
  if not Q.eof then Result := Q.Fields[0].asString ;
  Ferme(Q) ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : verdier
Créé le ...... : 11/06/2002
Modifié le ... : 11/06/2002
Description .. : verification qu'il n'existe pas d'écritures correspondant au
Suite ........ : compte général dans des exo définitivement cloturés
Mots clefs ... :
*****************************************************************}
Function TOF_CHGNATUREGENE.ExisteMouvementsAnterieurs (TobDet : Tob) : Boolean ;
begin
Result:=False ;
If ExisteSql('Select E_General,E_Exercice from ecriture left join exercice on E_Exercice=EX_Exercice where EX_Etatcpta="CLO" and E_General="'+TobDet.getValue('G_General')+'"') then result:=True ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Nathalie Payrot
Créé le ...... : 07/11/2001
Modifié le ... :   /  /
Description .. : Traitement des comptes.
Mots clefs ... :
*****************************************************************}
procedure TOF_CHGNATUREGENE.TraitementCompte ;
var
  AuxiLettrable : Boolean;
  ValLettrage   : string;
  ValPointage   : string;
  ValSuiviTreso : string;
  MajEcr    : Boolean;
  TobCpt    : Tob;
  TVADefaut : string;
  NbEcr     : Integer;
  idxCpt    : Integer ;
  CptAuxi   : string;
  ModePaie  : string;
begin
  NbEcr := 0;
  // récupération données du compte auxiliaire si renseigné
  CptAuxi := GetControlText('NEWAUXI') ;
  RechTiers (CptAuxi, AuxiLettrable, ModePaie);
  if AuxiLettrable then ModePaie := RechModePaie (ModePaie) ;
  TVADefaut := '' ;
  If (NewNature.Value='CHA') or ( NewNature.Value='PRO') then TVADefaut := RechTVA ;

  {On boucle sur les comptes généraux}
  for idxCpt := 0 to TobGene.Detail.Count-1 do begin
    TobCpt := TobGene.Detail[idxCpt];
    try
      ValLettrage   := '' ;
      ValSuiviTreso := TobCpt.getValue('G_SUIVITRESO') ;
      {JP 04/10/07 : en fait la gestion du pointage est inutile car si le compte divers est pointable,
                     une fois transformé en compte de banque, il le restera et inversement.
                     le seul cas qui peut poser un problème, est de transformer un divers pointable en un compte
                     qui ne l'est pas : mais dans ce cas là, la plupart des champs seront mis à jour avec les
                     information de lettrage, seuls resteront les champs e_datepointage et e_refpointage dans
                     la base => on met ValPointage à "-", pour les vider dans ces cas là}
      if (TobCpt.GetString('G_POINTABLE') = 'X') and
         ((NewNature.Value <> 'BQE') or (NewNature.Value <> 'CAI') or
          (NewNature.Value <> 'EXT') or (NewNature.Value <> 'DIV')) then ValPointage := '-'
                                                                    else ValPointage := 'X';

      if (NewNature.Value = 'COC') then ValSuiviTreso := 'ENC';
      if (NewNature.Value = 'COF') then ValSuiviTreso := 'DEC';

      {Gestion des comptes qui à l'origine sont collectifs : BPY le 25/06/2004 => Fiche 12349}
      if (Copy(NatureOri.Value, 1, 2) = 'CO') then begin
        {Dans le cas où ils restent des comptes collectifs}
        if (Copy(NewNature.Value,1,2) = 'CO') then begin
          {test des auxiliaires affectés au colectif (il faut que les auxiliaire soient divers)}
          if (ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_COLLECTIF="' + TobCpt.GetValue('G_GENERAL') +
             '" AND T_NATUREAUXI<>"DIV"')) then begin
            PGIBox('Modification impossible : Il existe des comptes auxiliaires non divers rattachés à ce collectif ('+ TobCpt.GetValue('G_GENERAL') + ').');
            Continue;
          end;

          {test des ecritures attachées au compte (il faut que les ecritures portent sur des comptes auxiliaires divers)}
          if (ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE IN (SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_GENERAL="' +
             TobCpt.GetValue('G_GENERAL') + '") AND T_NATUREAUXI<>"DIV"')) then begin
            PGIBox('Modification impossible : Il existe des écritures sur des comptes auxiliaires non divers rattachés à ce collectif ('+ TobCpt.GetValue('G_GENERAL') + ').');
            Continue;
          end;
        end

        {Il s'agit de comptes collectifs qui ne le seront plus après le traitement}
        else begin
          {Il ne faut pas qu'il y ait des comptes auxiliaires rattachés à ce collectifs}
          if (ExisteSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_COLLECTIF="' +
             TobCpt.GetValue('G_GENERAL') + '"')) then begin
            PGIBox('Modification impossible : Il existe des comptes auxiliaires rattachés à ce collectif ('+ TobCpt.GetValue('G_GENERAL') + ').');
            Continue;
          end;
        end;
      end;

      {On ne traite pas les comptes divers non lettrable que l'on souhaiterit changer en comptes de banques ou de caisses}
      MajEcr := not (((NatureOri.Value='DIV') and (TobCpt.GetValue('G_LETTRABLE') <> 'X')) and
                     ((NewNature.Value='BQE') or (NewNature.Value='CAI'))) ;

      if MajEcr then //gv 11/06/02
        If PGIASK(TraduireMemoire('Les mouvements portant sur le compte général '+TobCpt.GetValue('G_GENERAL')+' vont être modifiés. Confirmez-vous le traitement ?'),TraduireMemoire(Titre)) = MrNo then Exit; 

      {Il s'agit à l'origine d'un compte divers ...}
      if NatureOri.Value = 'DIV' then  begin                                                                                             // fiche 11825
        {... s'il est lettrable et que l'on le destine vers une nature non lettrable,
             il faudra annuler le lettrage sur les écriture}
        if (TobCpt.GetValue('G_LETTRABLE') = 'X') and
           ((NewNature.Value='BQE')or(NewNature.Value='CAI') or (NewNature.Value='IMO') or
           (NewNature.Value='CHA') or ( NewNature.Value='PRO')) then ValLettrage := '-';

        {... si l'on le destine vers un TIC ou TID, il faudra mettre à jour les infos de lettrage}
        if Copy(NewNature.Value,1,2) = 'TI' then begin
          ValLettrage := 'X' ;
          ModePaie := RechModePaie(GetParamsocSecur('SO_GCMODEREGLEDEFAUT',''));
        end;
      end

      {Changement de nature de comptes TID ou TIC}
      else begin
        {JP 19/10/05 : FQ 16797 : si le tiers est lettrable, le divers le sera aussi et l'on conservera
                       les infos de lettrage sur les écriture}
        if (NewNature.Value = 'DIV') and (TobCpt.GetValue('G_LETTRABLE') = 'X') then ValLettrage := 'X'
        else if Copy(NewNature.Value, 1, 2) <> 'CO' then ValLettrage := '-';
        { FQ 19085 BVE 12.04.07 }
        if (Copy(NewNature.Value, 1, 2) = 'CO') and (Copy(NatureOri.Value, 1, 2) = 'TI') then
        begin
          RechTiers (CptAuxi, AuxiLettrable, ModePaie);
          if ModePaie <> '' then
            ModePaie := RechModePaie (ModePaie)
          else
            ModePaie := RechModePaie (GetParamsocSecur('SO_GCMODEREGLEDEFAUT',''));
        end
        else
          ModePaie := '';
        { END FQ 19085 }
      end;

      {verification qu'il n'existe pas d'écritures correspondant au compte général dans des exo définitivement cloturés}
      if (((NewNature.Value='CHA') or ( NewNature.Value='PRO')) and
          (NatureOri.Value='DIV')) then if ExisteMouvementsAnterieurs(TobCpt) then Continue ; // gv 11/06/02

      BeginTrans ;
      // modification du compte
      ModifCompte ( TobCpt, ValLettrage, ValSuiviTreso, TVADefaut ) ;
      {modification des écritures si besoin}
      if MajEcr then
        NbEcr := NbEcr + ModifEcriture(TobCpt.GetValue('G_GENERAL'), CptAuxi, ValLettrage, ValPointage, ModePaie, AuxiLettrable);

      {Mise à jour du compte général}
      TobGene.Detail[idxCpt].UpdateDb ;

      CommitTrans;
    except
      Rollback;
      PgiBox(TraduireMemoire('ATTENTION : Programme Interrompu. Relancer le traitement.'),TraduireMemoire(Titre)) ;
      Exit;
    end;
  end;

  {mise à jour des cumuls debit credit du tiers
   JP 04/10/07 : Utilidation de la fonction de SoldeCPT}
  if CptAuxi <> '' then MajTotComptes([fbAux], False, False, '', CptAuxi);

  if NbEcr > 0 then
    PGIInfo(IntToStr(NbEcr) + TraduireMemoire(' écritures on été traitées.'), TraduireMemoire(Titre)) ;

  if (NewNature.Value='BQE') then
    PGIInfo(TraduireMemoire('N''oubliez pas d''aller renseigner les RIB sur les comptes modifiés,') + #13 +
            TraduireMemoire('ainsi que les journaux de banque.'), TraduireMemoire(Titre)) ;
end ;

{JP 26/08/04, FQ 14058 : Changement du nom de la procedure pour la mettre en accord avec le traitement}
procedure TOF_CHGNATUREGENE.BLancerOnClick (sender : TObject) ;
begin
  ListeCpt := Fiche.FListe ;
{$IFDEF EAGLCLIENT}
  QListeCpt := Fiche.Q.TQ;
{$ELSE}
  QListeCpt := Fiche.Q;
{$ENDIF}

  {JP 01/06/05 : FQ 15364, on limite le test aux comptes collectifs}
  if (GetControlText('NEWNATCPT') = 'COF') or
     (GetControlText('NEWNATCPT') = 'COC') or
     (GetControlText('NEWNATCPT') = 'COS') or
     (GetControlText('NEWNATCPT') = 'COD') then begin
    //SG6 04.03.05 FQ 15364
    if not LookUpValueExist(THEdit(getcontrol('NEWAUXI', true))) then
    begin
      PGIInfo('Le compte auxiliaire n''est pas valide.');
      Exit;
    end;
  end;

  // vérification qu'au moins un compte est sélectionné
  if (ListeCpt.nbSelected = 0) and (not ListeCpt.AllSelected) then
    begin
    PGIInfo(TraduireMemoire('Vous n''avez sélectionné aucun compte. Traitement impossible.'), TraduireMemoire(Titre)) ;
    Exit ;
    end ;
  // la nouvelle nature doir être renseignée
  if NewNature.Value = '' then
    begin
    PGIInfo(TraduireMemoire('Vous devez renseigner une nouvelle nature.'), TraduireMemoire(Titre)) ;
    SetFocusControl('NEWNATCPT') ;
    Exit ;
    end ;
  // si la nature est de type collectif, le compte auxiliaire doit être renseigné
  if (Copy(NewNature.Value,1,2)='CO') and (GetControlText('NEWAUXI')='') and (not ((Copy(NatureOri.Value,1,2) = 'CO'))) then
    begin
    PGIInfo(TraduireMemoire('Pour une nature de compte collectif, vous devez saisir un compte auxiliaire.'), TraduireMemoire(Titre)) ;
    SetFocusControl('NEWAUXI') ;
    Exit ;
    end ;

  If PGIASK(TraduireMemoire('Confirmez-vous le traitement ?'),TraduireMemoire(Titre)) = MrNo then exit ;

  // si changement de la nature vers Banque, affichage du message d'info pour aller saisir les RIB
  if (NewNature.Value='BQE') then if PGIASK(TraduireMemoire('Vous devrez modifier ces comptes pour saisir un RIB. Confirmez-vous le traitement ?'),TraduireMemoire(Titre)) = MrNo then exit ;

  EnableControls(Fiche, False) ;
  // récupération des comptes sélectionnés
  RecupMultiSelect ;
  // lancement du traitement

  {JP 04/10/07 : (vu EC et OG) Pour la Tréso, on ne touche pas à E_TRESOSYNCHRO, mais message d'avertissement}
  if EstComptaTreso then
    if PGIAsk(TraduireMemoire('Ce traitement ne touchera pas à la Trésorerie.') + #13#13 +
              TraduireMemoire('Deux cas posent un problème :') + #13#13 +
              TraduireMemoire('- Le(s) compte(s) étai(en)t synchronisable(s) mais plus maintenant.') + #13 +
              TraduireMemoire('  Vous devrez supprimez manuellement les flux prévisionnels en Trésorerie.') + #13#13 +
              TraduireMemoire('- Le(s) compte(s) n''étai(en)t pas synchronisable(s) mais le devien(nen)t.') + #13 +
              TraduireMemoire('  L''antériorité des écritures ne sera pas synchronisable ;') + #13 +
              TraduireMemoire('  seules les écritures créées après le traitement seront synchronisables.') + #13#13 +
              TraduireMemoire('Souhaitez-vous abandonner ?')
              , Titre) = mrYes then Exit;

  TraitementCompte ;
  EnableControls(Fiche, True) ;

  // rafraîchissement de la liste
  TToolBarButton97(GetControl('BCHERCHE')).Click ;
end ;

procedure TOF_CHGNATUREGENE.FListeOnDblClick (sender : TObject) ;
begin
{$IFDEF EAGLCLIENT}
  QListeCpt := Fiche.Q.TQ;
{$ELSE}
  QListeCpt := Fiche.Q;
{$ENDIF}

  // visu du compte sur double clic
  {JP 21/08/07 : FQ 21261 : GetField plutôt que Q.FindField pour le cwas}
  FicheGene(QListeCpt,'',GetField('G_GENERAL'),taConsult,0) ;
end ;

procedure TOF_CHGNATUREGENE.OnArgument (S : String ) ;
begin
  Inherited ;

  Ecran.HelpContext:=999999104;

  Fiche := TFMul(Ecran) ;
  ListeCpt := Fiche.FListe ;
{$IFDEF EAGLCLIENT}
  QListeCpt := Fiche.Q.TQ;
{$ELSE}
  QListeCpt := Fiche.Q;
{$ENDIF}
  NatureOri := THValComboBox(GetControl('G_NATUREGENE')) ;
  NewNature := THValComboBox(GetControl('NEWNATCPT')) ;
  Pages := TpageControl(GetControl('PAGES')) ;

//  NatureOri.Plus := 'AND (CO_CODE="DIV" OR CO_CODE LIKE "TI%") ' ; // bpy le 12/08/2003 fiche bug n° 12349
{ CA - 10/09/2004 - TROP DANGEREUX D'OUVRIR TOUTES LES NATURES : ON REDUIT LES POSSIBILITES COMME EN VERSION 5.00 }
  NatureOri.Plus := 'AND (CO_CODE="DIV" OR CO_CODE LIKE "TI%") ';

  NatureOri.ItemIndex := 0 ;
{ BPY 29/01/2003 Correction de la fiche 10541 }
  NewNature.Plus := 'AND (CO_CODE<>"EXT") AND (CO_CODE<>"DIV")' ;
{
  NewNature.Plus := 'AND (CO_CODE<>"IMO") AND (CO_CODE<>"EXT") AND (CO_CODE<>"DIV")' ;
}
{ Fin BPY }
  NatureOri.OnChange := NatureOriOnChange ;
  NewNature.OnChange := NewNatureOnChange ;


  {JP 26/08/04, FQ 14058 : BOuvrir ouvre la fiche sur la ligne en cours et BValider lance le traitement}
  TToolBarButton97(GetControl('BVALIDER')).OnClick := BLancerOnClick ;
  TToolBarButton97(GetControl('BOUVRIR')).OnClick := FListeOnDblClick ;

  TobGene := Tob.Create('MesComptes', Nil, -1 ) ;
{$IFDEF EAGLCLIENT}
  THGrid(GetControl('FLISTE')).OnDblClick := FListeOnDblClick ;
{$ELSE}
  THDBGrid(GetControl('FLISTE')).OnDblClick := FListeOnDblClick ;
{$ENDIF}

  Titre := 'Modification de nature des comptes généraux' ;
end ;

procedure TOF_CHGNATUREGENE.OnClose ;
begin
  TobGene.Free ;
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CHGNATUREGENE ] ) ;
end.
