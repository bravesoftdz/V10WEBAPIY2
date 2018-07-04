{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 25/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : HISTOGROUPEELEG ()
Mots clefs ... : TOF;HISTOGROUPEELEG
*****************************************************************}
{
PT1 02/07/2007 MF V_72 FQ 14480 Ajout boutons Exporter la liste et Voir légende Excel
                       sur fiche Acompte_GRP
PT2 02/08/2007 FC V_80 Rajout de la légende de la saisie groupée des éléments dynamiques et dossier
}
Unit UTofPGHistoGroupeeLegende ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul,

{$ENDIF}
     forms,
     uTob, 
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox,
     EntPaie,
     UTOF,
     UTofPGMulEltDynGroupee; //PT2 

Type
  TOF_PGHISTOGROUPEELEG = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    Procedure RendLibelle(Num : Integer;Var Libelle : String);

  end ;

Implementation

procedure TOF_PGHISTOGROUPEELEG.OnArgument (S : String ) ;
var Q : TQuery;
    T : Tob;
    GLegende : THGrid;
    i,Ident : Integer;
    Libelle : String;
    TobFille : Tob; // PT1
    NbElements : integer; //PT2
begin
  Inherited ;
// d PT1
  if (S = 'ACOMPTE_GRP') then
  begin

    GLegende := THGrid(GetControl('GLEGENDE'));
    T := Tob.Create('LesChamps',Nil,-1);
    TobFille := TOB.Create('', T, -1);
    TobFille.AddChampSupValeur('NOM', 'Salarié', FALSE);
    TobFille.AddChampSupValeur('LIBELLE', 'Matricule', FALSE);
    TobFille := TOB.Create('', T, -1);
    TobFille.AddChampSupValeur('NOM', 'Date de saisie', FALSE);
    TobFille.AddChampSupValeur('LIBELLE', 'Date_de_saisie', FALSE);
    TobFille := TOB.Create('', T, -1);
    TobFille.AddChampSupValeur('NOM', 'Libellé', FALSE);
    TobFille.AddChampSupValeur('LIBELLE', 'Libellé', FALSE);
    TobFille := TOB.Create('', T, -1);
    TobFille.AddChampSupValeur('NOM', 'A reporter', FALSE);
    TobFille.AddChampSupValeur('LIBELLE', 'A_reporter', FALSE);
    TobFille := TOB.Create('', T, -1);
    TobFille.AddChampSupValeur('NOM', 'Montant', FALSE);
    TobFille.AddChampSupValeur('LIBELLE', 'Montant', FALSE);

    T.PutGridDetail(GLegende,False,False,'NOM;LIBELLE');
    T.Free;

  end
  //DEB PT2
  else if (S = 'ELTDYN_GRP') then
  begin
    GLegende := THGrid(GetControl('GLEGENDE'));
    T := Tob.Create('LesChamps',Nil,-1);
    NbElements := PGTobLesElements.detail.Count;
    For i := 0 to NbElements - 1 do
    begin
      TobFille := TOB.Create('', T, -1);
      if (PGTypElement = 'ELD') then
        TobFille.AddChampSupValeur('NOM', '_' + PGTobLesElements.detail[i].GetValue('PGINFOSMODIF'), FALSE)
      else
        TobFille.AddChampSupValeur('NOM', '_' + PGTobLesElements.detail[i].GetValue('CODEELT'), FALSE);
      TobFille.AddChampSupValeur('LIBELLE', PGTobLesElements.detail[i].GetValue('LIBELLE'), FALSE);
    end;
    T.PutGridDetail(GLegende,False,False,'NOM;LIBELLE');
    T.Free;
  end
  //FIN PT2
  else
  begin
// f PT1
  GLegende := THGrid(GetControl('GLEGENDE'));
  Q := OpenSQL('SELECT PAI_SUFFIX,PAI_COLONNE,PAI_LIBELLE,PAI_IDENT FROM PAIEPARIM WHERE PAI_PREFIX="PSA"',True);
  T := Tob.Create('LesChamps',Nil,-1);
  T.LoadDetailDB('LesChamps','','',Q,False);
  Ferme(Q);
  For i := 0 to T.Detail.Count - 1 do
  begin
       Ident := T.Detail[i].GetValue('PAI_IDENT');
       Libelle := T.Detail[i].GetValue('PAI_LIBELLE');
       RendLibelle(Ident,Libelle);
       T.Detail[i].PutValue('PAI_LIBELLE',Libelle);
  end;
  T.PutGridDetail(GLegende,False,False,'PAI_COLONNE;PAI_LIBELLE');
  T.Free;
  end; //PT1
end ;

Procedure TOF_PGHISTOGROUPEELEG.RendLibelle(Num : Integer;Var Libelle : String);
begin
     If (Num >=110) and (Num <= 114) then //Salaire mensuel
     begin
          If (Num = 110) and (VH_Paie.PgNbSalLib > 0) then Libelle := 'Salaire mensuel : '+VH_Paie.PgSalLib1
          else if (Num = 111) and (VH_Paie.PgNbSalLib > 1) then Libelle := 'Salaire mensuel : '+VH_Paie.PgSalLib2
          else If (Num = 112) and (VH_Paie.PgNbSalLib > 2) then Libelle := 'Salaire mensuel : '+ VH_Paie.PgSalLib3
          else If (Num = 113) and (VH_Paie.PgNbSalLib > 3) then Libelle := 'Salaire mensuel : '+VH_Paie.PgSalLib4
          else If (Num = 114) and (VH_Paie.PgNbSalLib > 4) then Libelle := 'Salaire mensuel : '+VH_Paie.PgSalLib5
          else Libelle := '';
     end
     else If (Num >=115) and (Num <= 119) then //Salaire ANNUEL
     begin
          If (Num = 115) and (VH_Paie.PgNbSalLib > 0) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib1
          else if (Num = 116) and (VH_Paie.PgNbSalLib > 1) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib2
          else If (Num = 117) and (VH_Paie.PgNbSalLib > 2) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib3
          else If (Num = 118) and (VH_Paie.PgNbSalLib > 3) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib4
          else If (Num = 119) and (VH_Paie.PgNbSalLib > 4) then Libelle := 'Salaire annuel : '+VH_Paie.PgSalLib5
          else Libelle := '';
     end
     else If (Num >=157) and (Num <= 160) then // Champs organisation
     begin
          If (Num = 157) and (VH_Paie.PGNbreStatOrg > 0) then Libelle := VH_Paie.PGLibelleOrgStat1
          else if (Num = 158) and (VH_Paie.PGNbreStatOrg > 1) then Libelle := VH_Paie.PGLibelleOrgStat2
          else If (Num = 159) and (VH_Paie.PGNbreStatOrg > 2) then Libelle := VH_Paie.PGLibelleOrgStat3
          else If (Num = 160) and (VH_Paie.PGNbreStatOrg > 3) then Libelle := VH_Paie.PGLibelleOrgStat4
          else Libelle := '';
     end
     else If Num = 161 then //Code statistique
     begin
          If VH_Paie.PGLibCodeStat <> '' then Libelle := VH_Paie.PGLibCodeStat
          else Libelle := '';
     end
     else If (Num >=345) and (Num <= 348) then // Combos libre
     begin
          If (Num = 345) and (VH_Paie.PgNbCombo > 0) then Libelle := VH_Paie.PgLibCombo1
          else if (Num = 346) and (VH_Paie.PgNbCombo > 1) then Libelle := VH_Paie.PgLibCombo2
          else If (Num = 347) and (VH_Paie.PgNbCombo > 2) then Libelle := VH_Paie.PgLibCombo3
          else If (Num = 348) and (VH_Paie.PgNbCombo > 3) then Libelle := VH_Paie.PgLibCombo4
          else Libelle := '';
     end
     else If Num = 49 then Libelle := 'Libelle emploi'
     else If Num = 1 then Libelle := 'Matricule salarié';
end;


Initialization
  registerclasses ( [ TOF_PGHISTOGROUPEELEG ] ) ;
end.

