{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. : Unité de définition des procédures et fonctions permettant
Suite ........ : l'intégration  des IJSS brutes et nettes dans le bulletin.
Mots clefs ... : PAIE;PGBULLETIN;IJSS
*****************************************************************}
{
PT1   30/11/2004 MF V_60 Correction traitement maintien qd champ catégorie renseigné.
PT2   29/12/2004 MF V_60 correction erreur SQL DB2
PT3   12/01/2005 MF V_60 Modification de la recherche de la table critère à utiliser
PT4   28/01/2005 MF V_60 On n'intègre pas de règlement d'IJSS ne correspondant à aucun maintien
PT5   08/02/2005 MF V_60 Traitement des commentaires des règlements d'IJSS
PT6   29/04/2005 MF V_602 FQ 12234 Nvelle focntion pour annulation des commentaires d'IJSS
PT7   20/02/2006 MF V_65 Correction memcheck + warning
PT8   16/03/2006 MF V_65 FQ 12968 : PCM_VALCATEG remplacé par PCM_VALCATEGORIE
PT9   01/08/2006 MF V_70 correction :
                         - quand plusieurs lignes de règlement à intégrer et aucun
                           maintien correspondant ==> "Indice de liste hors limites"
PT10  05/06/2008 MF V_82  FQ 459 et FQ 15466 :                      

}
unit PGIJSS;

interface
uses
  HCtrls,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
{$ENDIF}
  EntPaie,
  Hent1,
//unused  HMsgBox,
//  PGVisuObjet,
  P5Util,
  SysUtils,
  UTOB;

//PT4 function  RecupereIJSS(Datef:tdatetime;Salarie,Action:string; var TOB_IJSS : TOB):boolean;
function RecupereIJSS(Datef:tdatetime;Salarie,Action:string; var TOB_IJSS :TOB; const Tob_Maintien, Tob_Etab : TOB):boolean;
procedure IntegreIJSSDansPaye(Tob_Rub, Salarie: tob; const DateD, DateF: tdatetime; const Action: string; const ChampCateg : string; Var TOB_IJSS : TOB; var RubIJNettes : string;const Subrogation : boolean);
procedure MajDateIntegr(var t: tob; const datef: tdatetime);
procedure RechercheRubIJSS(const TypeAbs: string; var RubIJBrutes, RubIJNettes, Aliment : string; const ChampCateg : string);
procedure EnleveCommIJss(Salarie, Tob_rub: tob; const Arub, Natrub : string; const dated, datef: tdatetime);  // PT6
//procedure SalEcritIJSS(var tob_IJSS: tob);

var
  Anciennete                            : WORD;
  Convention                            : string;
  Categorie                             : string;
implementation
//PT4 function RecupereIJSS(Datef:tdatetime;Salarie,Action:string; var TOB_IJSS :TOB):boolean;
function RecupereIJSS(Datef:tdatetime;Salarie,Action:string; var TOB_IJSS :TOB; const Tob_Maintien , Tob_Etab: TOB):boolean;

var
  st                            : string;
  t                             : Tob;
  Q                             : TQuery;
//  MaDate                        : TDateTime ;
  tMaintien, Treglt             : TOB;      // PT4
  Wi, i, j                      : integer;  // PT4
  TotMaintien                   : integer;  // PT4
begin
  result:=False;
  if Action = 'M' then
    st:='SELECT * FROM REGLTIJSS WHERE'+
        ' PRI_SALARIE = "'+   Salarie+ '"'+
        ' AND (PRI_DATEREGLT <= "' + usdatetime(Datef)+'"' +
        ' AND (PRI_DATEINTEGR <= "' + usdatetime(10)+'"'+
        ' OR PRI_DATEINTEGR="'+usdatetime(Datef)+'"))'
  else
    st:='SELECT * FROM REGLTIJSS WHERE'+
        ' PRI_SALARIE = "'+   Salarie+ '"'+
        ' AND PRI_DATEREGLT <= "' + usdatetime(Datef)+'"' +
        ' AND PRI_DATEINTEGR <= "' + usdatetime(10)+'"';
  Q:=OpenSql(st, TRUE);
  if not Q.eof then
    begin
    Tob_IJSS:= Tob.create('les règlements d''IJSS',nil,-1);
    Tob_IJSS.LoadDetailDB('REGLTIJSS','','',Q,False);
//    MaDate := Idate1900;
    { Si règlement non intégré existe renvoie true pour réintégration paye }
//    if Tob_IJSS.Findfirst(['PRI_DATEINTEGR'],[MaDate],false) <> nil then
      result := true;
    end
  else
    if Tob_IJSS<>nil then FreeAndNil (Tob_IJSS);
  Ferme(Q);

  if (Action = 'M') and (Tob_IJSS<>nil) and (result) then
    { On réintègre tous les règlement même ceux déjà intégré }
    begin
    t:= Tob_IJSS.Findfirst([''],[''],false);
    while t<> nil do
      begin
      T.putvalue('PRI_DATEINTEGR',0);
      T:= Tob_IJSS.Findnext([''],[''],false)
      end;
    end;
//d PT4
     if (VH_Paie.Pgmaintien) and (VH_Paie.PgGestIJSS) and
        (Tob_IJSS <> nil) and (Tob_IJSS.Detail.count > 0) then
     begin
       Totmaintien := 0 ;
       Wi := 0;
       for i := 0 to Tob_IJSS.detail.count-1 do
       begin
         TReglt := Tob_IJSS.detail[wi];
// PT9         Wi := i ;
// Y-a-t'il eu du maintien sur l'absence dont les IJSS on été réglées ?
// si non on ne doit pas intégrer les retenues IJSS s'il n'y a pas de subrogation
         st := 'SELECT * FROM MAINTIEN WHERE '+
               'PMT_SALARIE = "'+treglt.getValue('PRI_SALARIE')+'" AND '+
               'PMT_DATEDEBUTABS <= "'+
               UsDateTime(treglt.getValue('PRI_DATEFINABS'))+'" AND '+
               'PMT_DATEFINABS >= "'+
               UsDateTime(treglt.getValue('PRI_DATEDEBUTABS'))+'" ';
         Q := opensql(st, True);
         TotMaintien := 0;

         if not Q.eof then
         begin
           Q.first;
           While not Q.eof do
           begin
             if (Q.FindField('PMT_MTMAINTIEN').AsFloat <> 0) then
               TotMaintien := TotMaintien + 1;
             Q.Next;
           end;
         end;
         if (Tob_Maintien <> nil) then
         for j:= 0 to Tob_Maintien.detail.count-1 do
         begin
           tMaintien := Tob_Maintien.Detail[j];
           if (tmaintien.GetValue ('PMT_SALARIE') = treglt.getValue('PRI_SALARIE')) and
              (tmaintien.GetValue ('PMT_DATEDEBUTABS') <= treglt.getValue('PRI_DATEFINABS')) and
              (tmaintien.GetValue ('PMT_DATEFINABS') >= treglt.getValue('PRI_DATEDEBUTABS')) then
             if (tmaintien.GetValue ('PMT_PCTMAINTIEN') <> 0) then
               TotMaintien := TotMaintien + 1;
         end;
         if (Tob_Etab <> nil) and (Tob_etab.GetValue('ETB_SUBROGATION') <> 'X') and
            (TotMaintien = 0) then
         begin
           TReglt.free;
// d PT9
//            Wi := Wi -1;
//           if (Wi < 0) then Wi := 0;
         end
         else
           Wi := Wi + 1;
// f PT9           
       end;
     end;
//f PT4

end;
procedure IntegreIJSSDansPaye(Tob_Rub, Salarie: tob; const DateD, DateF: tdatetime; const Action: string; const ChampCateg : string; var TOB_IJSS :TOB; var RubIJNettes : string;const Subrogation :boolean);
var
  Q                                                     : TQuery;
  t                                                     : tob;
  TypeAbs, TypeAbsP, Aliment,  Natrub                   : string;
  RubIJBrutes,RubIJBrutesP, RubIJNettesP    : string;
  ValIJBrutes,ValIJNettes                               : double;
//d PT3
  RechDebutAbs                                          : boolean;
  DateDebutAbsReel                                      : TDateTime;
  DateFinAbsPrec                                        : TDateTime;
  st                                                    : string;
// f PT3
  i,j             : integer; // PT5
begin
  if BullCompl = 'X' then exit;

  ValIJBrutes := 0;
  ValIJNettes := 0;
  i := 1; // PT5
  j := 1; // PT5
  if TOB_IJSS = nil then exit;

  t := Tob_IJSS.findfirst([''], [''], false);
  if t = nil then exit;
  TypeAbs := t.getvalue('PRI_TYPEABS');
  TypeAbsP := TypeAbs;
// d PT3
// recherche date début absence correspondant au règlement
  RechDebutAbs := True;
  DateDebutAbsReel := t.getValue('PRI_DATEDEBUTABS');
  DateFinAbsPrec := DateDebutAbsReel-1 ;
  while RechDebutAbs do
  begin
    st := 'SELECT PCN_DATEDEBUTABS FROM ABSENCESALARIE WHERE'+
          ' PCN_SALARIE = "'+   Salarie.GetValeur(iPSA_SALARIE)+ '"'+
          ' AND PCN_GESTIONIJSS ="X"'+
          ' AND PCN_DATEFINABS ="'+UsDateTime(DateFinAbsPrec)+'"'+
          ' AND PCN_TYPECONGE IN (SELECT PMA_MOTIFABSENCE FROM MOTIFABSENCE'+
          ' WHERE ##PMA_PREDEFINI## AND PMA_TYPEABS="'+TypeAbs+'")';
    Q := opensql(st, True);
    if not Q.eof then
    begin
      DateDebutAbsReel := Q.Fields[0].AsDateTime;
      DateFinAbsPrec := DateDebutAbsReel-1;
    end
    else
      RechDebutAbs := False;
    ferme (Q);
  end;

//  Anciennete := AncienneteMois (Salarie.GetValeur(iPSA_DATEANCIENNETE), t.getvalue('PRI_DATEDEBUTABS'));
  Anciennete := AncienneteMois (Salarie.GetValeur(iPSA_DATEANCIENNETE), PlusDate(DateDebutAbsReel,1,'J'));

// f PT3
  Convention :=  Salarie.GetValeur(iPSA_CONVENTION);

  if (ChampCateg <> '') then
  Categorie := Salarie.GetValue(ChampCateg);

  if (Convention='') then
   Convention := '000';

  RechercheRubIJSS (TypeAbs, RubIJBrutes, RubIJNettes, Aliment,ChampCateg);
  RubIJBrutesP := RubIJBrutes;
  RubIJNettesP := RubIJNettes;

  Natrub := 'AAA';

  if (Action = 'M') then
  begin
    EnleveCommIJSS(Salarie, Tob_rub, RubIJBrutes, Natrub, DateD, Datef);      // PT6
    EnleveCommIJSS(Salarie, Tob_rub, RubIJNettes, Natrub, DateD, Datef);      // PT6
  end;

  while t <> nil do
  begin
{PT10
    TypeAbs := t.getvalue('PRI_TYPEABS');
    if (TypeAbs <> TypeAbsP) and
      ((RubIJBrutesP <> RubIJBrutes) or (RubIJNettesP <> RubIJNettes)) then}
    if (TypeAbs <> TypeAbsP) then
    begin
      if (RubIJBrutesP <> RubIJBrutes) then
      begin
{PT10
        IntegreRub(tob_rub, Salarie, typeAbsP, Aliment, DateD, DateF, ValIJBrutes, Natrub, RubIJBrutes,True,False);}
        IntegreRub(tob_rub, Salarie, typeAbsP, Aliment, DateD, DateF, ValIJBrutes, Natrub, RubIJBrutesP,True,False);
        ValIJBrutes := 0;
      end;

      if (RubIJNettesP <> RubIJNettes) then
      begin
        if (Subrogation) then // Subrogation = True  si on traite uniquement les IJSS
{PT10
          IntegreRub(tob_rub, Salarie, typeAbsP, Aliment, DateD, DateF, ValIJNettes, Natrub, RubIJNettes, True,False);}
          IntegreRub(tob_rub, Salarie, typeAbsP, Aliment, DateD, DateF, ValIJNettes, Natrub, RubIJNettesP, True,False);
        ValIJNettes := 0;
      end;

      typeAbsP := TypeAbs;

      RechercheRubIJSS (TypeAbs, RubIJBrutes, RubIJNettes, Aliment,ChampCateg);

      if RubIJBrutes <> RubIJBrutesP then
      begin
        EnleveCommIJSS(Salarie, Tob_rub, RubIJBrutes, Natrub, DateD, Datef);  // PT6
        RubIJBrutesP := RubIJBrutes;
      end;
      if (RubIJNettes <> RubIJNettesP) and (Subrogation) then //PT1
      begin
        EnleveCommIJSS(Salarie, Tob_rub, RubIJNettes, Natrub, DateD, Datef);   // PT6
        RubIJNettesP := RubIJNettes;
      end;
    end;
//d PT5
    EcritCommIJ(Tob_rub, Salarie, t, i, DateD, Datef, RubIJBrutes, NatRub);
    if (Subrogation) then
      EcritCommIJ(Tob_rub, Salarie, t, j, DateD, Datef, RubIJNettes, NatRub);
//f PT5
    ValIJBrutes := ValIJBrutes + t.GetValue('PRI_MTIJSSBRUTES');
    ValIJNettes := ValIJNettes + t.GetValue('PRI_MTREGLE');
    MajDateIntegr(t, DateF);
    t := Tob_IJSS.findnext([''], [''], false);
// d PT10
   if (t <> nil) then
    begin
      TypeAbs := t.getvalue('PRI_TYPEABS');
      RubIJBrutesP := RubIJBrutes;
      RubIJNettesP := RubIJNettes;
      RechercheRubIJSS (TypeAbs, RubIJBrutes, RubIJNettes, Aliment,ChampCateg);
    end;
// f PT10
  end;

    // Ecriture du dernier
  IntegreRub(tob_rub, Salarie, typeAbsP, ALiment, DateD, DateF, ValIJBrutes, Natrub, RubIJBrutes, True,False);
  if (Subrogation) then // Subrogation = True  si on traite uniquement les IJSS
    IntegreRub(tob_rub, Salarie, typeAbsP, ALiment, DateD, DateF, ValIJNettes, Natrub, RubIJNettes, True,False);

end;
procedure RechercheRubIJSS(const TypeAbs: string; var RubIJBrutes, RubIJNettes, Aliment: string;const ChampCateg : string);
var
  Q: tQuery;
  st: string;
// d PT3
  i, Poids, PoidsMax, NoCrit                            : Integer;
  Tob_CritMaintien, tCritmaintien                       : TOB;
// f PT3
begin
  NoCrit := 0; // PT7
  Aliment := '';
  RubIJBrutes := '';
  RubIJNettes := '';
//d PT3
  st := 'SELECT * FROM CRITMAINTIEN WHERE ##PCM_PREDEFINI## (PCM_TYPEABS = "' +
        TypeAbs+'" OR PCM_TYPEABS="") AND  PCM_BORNEFINANC >= '+
        IntToStr(Anciennete)+' AND PCM_BORNEDEBANC <= '+IntToStr(Anciennete);

  if (Convention <> '') then
    st := st + ' AND (PCM_CONVENTION ="'+Convention+'" OR '+
          'PCM_CONVENTION ="000")'
  else
    st := st + ' AND PCM_CONVENTION ="000"';

  if (ChampCateg<>'') then
    if (Categorie = '') then
      st := st + ' AND (PCM_VALCATEGORIE = "" OR PCM_VALCATEGORIE="<<Tous>>")' // PT8
    else
      st := st + ' AND (PCM_VALCATEGORIE LIKE "%'+Categorie+'%"'+      // PT8
                 ' OR PCM_VALCATEGORIE="" OR PCM_VALCATEGORIE="<<Tous>>")';   // PT8

  Q:=OpenSql(st, TRUE);

  if not Q.eof then
  begin
    PoidsMax := 0;
    Tob_CritMaintien := Tob.create('les tables de critère ',nil,-1);
    Tob_CritMaintien.LoadDetailDB('CRITMAINTIEN','','',Q,False);

    // plusieurs Critères de maintien peuvent correspondre à la sélection
    // il faut définir l'élément de la table CRITMAINTIEN optimum, on donne
    // donc à chaque élément de la tob un poids
    if (Tob_CritMaintien <> nil) and (Tob_CritMaintien.Detail.count <> 0) then
    begin
      tCritMaintien := Tob_CritMaintien.FindFirst([''],[''],false);
      for i:= 1 to Tob_CritMaintien.Detail.count do
      begin
        Poids := 0;
        // PREDEFINI
        if (tCritMaintien.GetValue('PCM_PREDEFINI') = 'CEG') then
          // CEG
          Poids := Poids + 1
          else
            if (tCritMaintien.GetValue('PCM_PREDEFINI') = 'STD') then
              // STD
              Poids := Poids + 2
            else
              // DOS
               Poids := Poids + 3;

        // TYPEABS
        if (tCritMaintien.GetValue('PCM_TYPEABS') = TypeAbs) then
          Poids := Poids + 2
        else
          Poids := Poids +1;

        // CONVENTION
        if ((tCritMaintien.GetValue('PCM_CONVENTION') = Convention) or
            ((tCritMaintien.GetValue('PCM_CONVENTION') = '000') and
             (Convention = ''))) then
           Poids := Poids + 2
        else
           Poids := Poids +1 ;

        // VALCATEGORIE
        if ((tCritMaintien.GetValue('PCM_VALCATEGORIE') = Categorie) or      // PT8
            ((tCritMaintien.GetValue('PCM_VALCATEGORIE') = '<<Tous>>') and   // PT8
            (Categorie=''))) then
           Poids := Poids + 2
        else
           Poids := Poids +1;

        TCritmaintien.AddChampSupValeur('POIDS', Poids, False);
        if (Poids > PoidsMax) then
        // c'est lélément de poids le + élévé qui sera retenu
        begin
          PoidsMax := Poids;
          NoCrit := i;
        end;
        tCritMaintien := Tob_CritMaintien.FindNext([''],[''],false);
      end;
    end;

    tCritMaintien := Tob_CritMaintien.Detail[NoCrit-1];

    Aliment := 'MON';     // Montant
    RubIJBrutes := tCritMaintien.GetValue('PCM_RUBIJSSBRUTE');
    RubIJNettes := tCritMaintien.GetValue('PCM_RUBIJSSNETTE');
  end;
//f PT3
  Ferme(Q);
  FreeAndNil (Tob_CritMaintien);                              // PT7 memcheck
end; { fin RechercheRubIJSS}
// d PT6
procedure EnleveCommIJss(Salarie, Tob_rub: tob; const Arub, Natrub : string; const dated, datef: tdatetime);
var
  t: tob;
begin
  if BullCompl = 'X' then exit;

  // on vire les libelles de la tob_rub créés précédemment
  T := Tob_Rub.FindFirst(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
    [Natrub, DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  while T <> nil do
  begin
    if ((copy(T.GetValue('PHB_RUBRIQUE'), 1, length(ARub) + 1) = ARub + '.') and
     (T.GetValue('PHB_ORIGINELIGNE') = 'IJS')) then
      T.free;
    T := Tob_Rub.FindNext(['PHB_NATURERUB', 'PHB_DATEDEBUT', 'PHB_DATEFIN', 'PHB_SALARIE'],
      [Natrub, DateD, DateF, Salarie.GetValeur(iPSA_SALARIE)], TRUE);
  end;
end;
// f PT6

procedure MajDateIntegr(var t: tob; const datef: tdatetime);
begin
  if T = nil then exit;
  T.putvalue('PRI_DATEINTEGR', datef);
end;  { fin MajDateIntegr}

end.
