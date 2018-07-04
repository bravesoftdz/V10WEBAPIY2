{***********UNITE*************************************************
Auteur  ...... : FLO
Créé le ...... : 06/07/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TOTALCOMPTEURPRES ()
Mots clefs ... : TOF;TOTALCOMPTEURPRES
*****************************************************************}
Unit UTOFPGTotalCompteurPres;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     forms,
     sysutils, 
     ComCtrls,
     UTOF ;

Type
  TOF_PGTOTALCOMPTEURPRES = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  private
    Salarie, StWhereSalarie, StWhereGlobal  : String;
    DateDeb, DateFin : TDateTime;

    procedure ChargeDonnees;
    procedure SelectionFiltre (Sender : TObject);
  end ;

Implementation

Uses
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     UTOB,
{$ELSE}
     uTob,
{$ENDIF}
     HCtrls;

procedure TOF_PGTOTALCOMPTEURPRES.OnArgument (S : String) ;
Var Action  : String;
begin
  Inherited ;

     // Récupération des paramètres en entrée
     Action  := ReadTokenSt(S);
     Salarie := ReadTokenSt(S);
     DateDeb := StrToDate(ReadTokenSt(S));
     DateFin := StrToDate(ReadTokenSt(S));
     StWhereSalarie := ReadTokenSt(S);
     StWhereGlobal  := ReadTokenSt(S);

     // Mise à jour du libellé de la fenêtre
     If Salarie <> '' Then
          SetControlText('TITRE', 'Salarié :  '+Salarie+'   '+RechDom('PGSALARIE', Salarie, False))
     Else
          SetControlText('TITRE', 'Ensemble des salariés');


     // Mise à jour de la période concernée
     SetControlText('LBLDATEDEBUT', DateToStr(DateDeb));
     SetControlText('LBLDATEFIN', DateToStr(DateFin));

     SetControlEnabled('RBFILTRE', True);
     SetControlEnabled('RBFILTRE_', True);

     (GetControl('RBFILTRE') As THRadioButton).OnClick := SelectionFiltre;
     (GetControl('RBFILTRE_') As THRadioButton).OnClick := SelectionFiltre;

     // Chargement des données totales
     ChargeDonnees;
end ;

procedure TOF_PGTOTALCOMPTEURPRES.ChargeDonnees;
Var TobCompteurs : TOB;
    StSQL        : String;
    Grille       : THGrid;
    Filtre       : String;
    i            : Integer;
begin
     If (GetControl('RBFILTRE') As THRadioButton).Checked Then
       Filtre := 'PYR_LIBELLE'
     Else
       Filtre := 'PYP_THEMEPRE';

     // Création de la requête
     StSQL := 'SELECT '+Filtre+', SUM(PYP_QUANTITEPRES) AS QUANTITE FROM PGPRESENCESAL WHERE PYP_DATEDEBUTPRES>="'+UsDateTime(DateDeb)+
              '" AND PYP_DATEFINPRES<="'+UsDateTime(DateFin)+'" AND PYR_LIBELLE IN (SELECT PYR_LIBELLE FROM COMPTEURPRESENCE WHERE PYR_EDITPLANPRES="X"';
     If StWhereSalarie <> '' Then StSQL := StSQL + ' AND ' + StWhereSalarie;
     StSQL := StSQL + ')';
     If Salarie <> '' Then StSQL := StSQL + ' AND PYP_SALARIE="'+Salarie+'" ';
     If (StWhereGlobal <> '') Then StSQL := StSQL + ' AND ' + StWhereGlobal;
     StSQL := StSQL + ' GROUP BY '+Filtre+' ORDER BY '+Filtre;

     // Chargement de la TOB
     TobCompteurs := TOB.Create('LesCompteurs', Nil, -1);
     TobCompteurs.LoadDetailFromSQL(StSQL);

     // Récupération du libellé de chaque thème
     If Filtre = 'PYP_THEMEPRE' Then
     Begin
          For i := 0 To TobCompteurs.Detail.Count - 1 Do
               TobCompteurs.Detail[i].AddChampSupValeur('PYR_LIBELLE', RechDom('PGTHEMECOMPTEURPRES', TobCompteurs.Detail[i].GetValue('PYP_THEMEPRE'), False));
     End;

     Grille := GetControl('COMPTEURS') As THGrid;

     // Réinitialisation de la grille avant de faire le PutGridDetail pour ne pas avoir de pb de rafraîchissement
     Grille.RowCount :=0;
     TobCompteurs.PutGridDetail(Grille,False, False, 'PYR_LIBELLE;QUANTITE');

     // Formatage de la grille
     Grille.ColTypes[0] := 'C';
     Grille.ColTypes[1] := 'R';

     Grille.ColFormats[0] := '';
     Grille.ColFormats[1] := '##0.00';

     Grille.ColAligns[0] := taLeftJustify;
     Grille.ColAligns[1] := taRightJustify;

     FreeAndNil (TobCompteurs);
end;

procedure TOF_PGTOTALCOMPTEURPRES.SelectionFiltre(Sender: TObject);
begin
     ChargeDonnees;
end;

Initialization
  registerclasses ( [ TOF_PGTOTALCOMPTEURPRES ] ) ;
end.
