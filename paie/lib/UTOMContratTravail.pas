{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 27/06/2001
Modifié le ... :   /  /
Description .. : TOM gestion des contrats de travail
Mots clefs ... : PAIE;CONTRATTRAVAIL
*****************************************************************}
{PT1 27/06/01 V536 PH pb valeur de la tablette PGTYPECONTRAT CDD ==> CCD test revu
PT2 06/12/01 V563 JL  ajout message si date début essai différent de date début de contrat ,+ Date renouvellement inférieur date fin essai
PT2 06/12/01 V563 JL  les contrôles sont effectués dans la procédure OnUpdate au lieu de OnChangeField
PT3 27/03/02 V571 JL  Gestion des champs pour intermittents du spectacle
PT4 06/09/02 V582 Ph  Gestion affichage 5éme elt de salaire
PT5 22/10/02 V585 Ph  test contrat de type CDD en fonction de CDD ou CCD
PT6 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
    09/10/2003 JL V_42 Initialisation des dates période essai dans OnUpdateRecord
PT7 27/06/2003 JL V_42 FQ : 10193 vérification chevauchement des dates de contrat
PT8 09/09/2003 JL V_42 Ajout établissement
PT9 09/10/2003 JL V_42 FQ 10193  : message bloquant si date début période essai < date début contra
PT10 09/05/2004 PH V_50 FQ 11244 : les champs DATE sont initialisés avec une valeur cohérente
PT11 22/04/2004 PH V_60 Caption de la forme
PT12 13/10/2005 PH V_65 FQ 11908 Recup de la Date entree et date de sortie pour date de début et de fin
PT13 28/02/2006 GGS V6.50 FQ 12310 Ajout Date de fin de travail effectif
PT14 20/06/2006 GGS V7.00 FQ 11732 Message d'information si CDD sans date de fin
PT15 02/10/2006 GGS V7.20 FQ 13340 Alimentation motif de sortie contrat si motif de sortie salarie renseigné
PT16 12/12/2006 PH V7.20 Correction affichage zones montant Utilisation d'une TOB au lieu d'une Query.
PT17 05/01/2007 GGS V8.00 Correction plantage (ajout du tob_create qui avait 'sauté'
PT18 11/12/2007 PH V8.00 FQ 14924 Récup des éléments par défaut du contrat précédant

}
unit UTOMCONTRATTRAVAIL;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Spin,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}HDB, DBCtrls, Fiche,
{$ELSE}
  eFiche,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97;

type
  TOM_CONTRATTRAVAIL = class(TOM)
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnNewRecord; override;
  private
    Salarie, LeTitre: string;
    TOB_Sal: TOB; // TOB du salarie traité
    procedure AccesCheckBoxIS(Sender: TObject); //PT3
  end;

implementation
{ TOM_CONTRATTRAVAIL }
uses Paramsoc, EntPaie, PgOutils, P5Util, P5def;

procedure TOM_CONTRATTRAVAIL.OnArgument(stArgument: string);
var
  num: Integer;
  numero: string;
  TLieu: THLabel;
{$IFNDEF EAGLCLIENT}
  Mensuel, Annuel: THDBEdit;
  TCHeure, TCCachet: TDBCheckBox;
{$ELSE}
  Mensuel, Annuel: THEdit;
  TCHeure, TCCachet: TCheckBox;
{$ENDIF}
begin
  inherited;
  Salarie := Trim(StArgument);
  Salarie := ReadTokenSt(Salarie); // Recup Etablissement sur lequel on travaille
  TOB_Sal := nil;
  LeTitre := 'Contrats de travail Paie & GRH';
  // Parametrage  Affichage des champs salaires
  for num := 1 to VH_Paie.PgNbSalLib do
  begin
    Numero := InttoStr(num);
    if Num > VH_Paie.PgNbSalLib then break;
{$IFNDEF EAGLCLIENT}
    Mensuel := THDbEdit(GetControl('PCI_SALAIREMOIS' + Numero));
    Annuel := THDbEdit(GetControl('PCI_SALAIRANN' + Numero));
{$ELSE}
    Mensuel := THEdit(GetControl('PCI_SALAIREMOIS' + Numero));
    Annuel := THEdit(GetControl('PCI_SALAIRANN' + Numero));
{$ENDIF}
    if Mensuel <> nil then Mensuel.Visible := TRUE;

    if Annuel <> nil then Annuel.Visible := TRUE;
    TLieu := THLabel(GetControl('SALLIB' + Numero));
    if (TLieu <> nil) then
    begin
      TLieu.Visible := TRUE;
      if Num = 1 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib1;
      end;
      if Num = 2 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib2;
      end;
      if Num = 3 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib3;
      end;
      if Num = 4 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib4;
      end;
      //PT4 06/09/02 V582 Ph  Gestion affichage 5éme elt de salaire
      if Num = 5 then
      begin
        TLieu.Caption := VH_Paie.PgSalLib5;
      end;
    end;
    if num = 1 then
    begin
      Tlieu := THLabel(GetControl('LSM'));
      if (TLieu <> nil) then TLieu.Visible := TRUE;
      Tlieu := THLabel(GetControl('LSA'));
      if (TLieu <> nil) then TLieu.Visible := TRUE;
    end;
  end;
{$IFNDEF EAGLCLIENT}
  TCHeure := TDBCheckBox(GetControl('PCI_ISHEURES')); //PT3
  TCCachet := TDBCheckBox(GetControl('PCI_ISCACHETS'));
{$ELSE}
  TCHeure := TCheckBox(GetControl('PCI_ISHEURES'));
  TCCachet := TCheckBox(GetControl('PCI_ISCACHETS'));
{$ENDIF}
  if TCHeure <> nil then TCHeure.OnClick := AccesCheckBoxIS;
  if TCCachet <> nil then TCCachet.OnClick := AccesCheckBoxIS;
end;

procedure TOM_CONTRATTRAVAIL.OnChangeField(F: TField);
var
  Fin, Debut: Variant;

begin
  inherited;
  if F.FieldName = ('PCI_TYPECONTRAT') then
  begin
    if GetField('PCI_TYPECONTRAT') <> 'CCD' then //PT1 V536
    begin
      SetControlEnabled('PCI_NONPRECARITE', FALSE);
      SetControlVisible('PCI_NONPRECARITE', FALSE);
      SetControlVisible('TPCI_NONPRECARITE', FALSE);
    end
    else
    begin
      SetControlEnabled('PCI_NONPRECARITE', TRUE);
      SetControlVisible('PCI_NONPRECARITE', TRUE);
      SetControlVisible('TPCI_NONPRECARITE', TRUE);
    end;
  end;

  if F.FieldName = ('PCI_FINCONTRAT') then
  begin
    Fin := GetField('PCI_FINCONTRAT');
    Debut := GetField('PCI_DEBUTCONTRAT');
    if (Fin > 10) and (Fin < Debut) then
    begin
      PGIBox('La date de fin de contrat est inférieure à la date de début', LeTitre);
      // PT6 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
      SetField('PCI_FINCONTRAT', IDate2099);
      SetFocusControl('PCI_DEBUTCONTRAT');
    end;
    SetField('PCI_FINTRAVAIL', Fin); //PT13
  end;
  {
  *************PT 3 Tous les contrôles sont maintenant effectués dans la procédure on update**************************

  if (F.FieldName='PCI_ESSAIFIN') OR (F.FieldName='PCI_ESSAIDEBUT') then
    begin
    Fin := GetField ('PCI_ESSAIFIN');
    Debut := GetField ('PCI_ESSAIDEBUT');
    if (Fin > 10) AND (Fin < Debut) then
     begin
     PGIBox ('La date de fin de période d''essai est inférieure à la date de début',LeTitre);
     SetField ('PCI_ESSAIFIN', NULL);
     SetFocusControl('PCI_ESSAIDEBUT');
     end;
    end;

  if (F.FieldName='PCI_RNVESSAIFIN') OR (F.FieldName='PCI_RNVESSAIDEB') then
    begin
    Fin := GetField ('PCI_RNVESSAIFIN');
    Debut := GetField ('PCI_RNVESSAIDEB');
    if (Fin > 10) AND (Fin < Debut) then
     begin
     PGIBox ('La date de fin de période d''essai est inférieure à la date de début',LeTitre);
     SetField ('PCI_RNVESSAIFIN', NULL);
     SetFocusControl('PCI_RNVESSAIDEB');
     end;
    If F.FieldName='PCI_RNVESSAIDEB' Then        //PT2
       begin
       Fin:=GetField('PCI_ESSAIFIN');
       If (Debut < Fin) and (Fin>10) and (debut>10) Then
          begin
          PGIBox('La date de début  du renouvellement de la période d''essai ne peut être inférieur à la date de fin de la période d''essai précédente',LeTitre);
          SetFocusControl('PCI_RNVESSAIDEB');
          end;
       end;
    end;
  // PT2
  if (F.FieldName='PCI_ESSAIDEBUT') then
     begin
     Debut:=GetField('PCI_ESSAIDEBUT');
     DebutCt:=GetField('PCI_DEBUTCONTRAT');
     If (Debut>10) and (Debut<>DebutCt) Then PGIBox('La date de début de la période d''essai est différente de la date de début du contrat',LeTitre);
     end;
  // FIN PT2   }
end;

procedure TOM_CONTRATTRAVAIL.OnLoadRecord;
var
  Q: TQuery;
  st, numero: string;
  num: Integer;
  T1, T2: TOB;
{$IFNDEF EAGLCLIENT}
  Mensuel, Annuel: THDBEdit;
{$ELSE}
  Mensuel, Annuel: THEdit;
{$ENDIF}
begin
  inherited;

  // PT5 22/10/02 V585 Ph  test contrat de type CDD en fonction de CDD ou CCD
  if (GetField('PCI_TYPECONTRAT') <> 'CDD') or (GetField('PCI_TYPECONTRAT') <> 'CCD') then
  begin
    SetControlEnabled('PCI_NONPRECARITE', FALSE);
    SetControlVisible('PCI_NONPRECARITE', FALSE);
    SetControlVisible('TPCI_NONPRECARITE', FALSE);
  end
  else
  begin
    SetControlEnabled('PCI_NONPRECARITE', TRUE);
    SetControlVisible('PCI_NONPRECARITE', TRUE);
    SetControlVisible('TPCI_NONPRECARITE', TRUE);
  end;

  if DS.State in [dsInsert] then
  begin // PT12
    St := 'SELECT PSA_LIBELLE,PSA_PRENOM,PSA_SALAIREMOIS1,PSA_SALAIREMOIS2,PSA_SALAIREMOIS3,PSA_SALAIREMOIS4,PSA_SALAIREMOIS5,PSA_SALAIRANN1,PSA_DATEENTREE,PSA_DATESORTIE,' +
      'PSA_MOTIFSORTIE,' +
      'PSA_SALAIRANN2,PSA_SALAIRANN3,PSA_SALAIRANN4,PSA_SALAIRANN5,PSA_HORAIREMOIS,PSA_HORHEBDO,PSA_HORANNUEL,PSA_TAUXHORAIRE,PSA_CONFIDENTIEL,PSA_LIBELLEEMPLOI' +
      ' FROM SALARIES WHERE PSA_SALARIE= "' + Salarie + '"'
  end
  else st := 'SELECT PSA_LIBELLE,PSA_PRENOM,PSA_CONFIDENTIEL FROM SALARIES WHERE PSA_SALARIE= "' + Salarie + '"';

  Q := OpenSQL(st, TRUE);
  T1 := TOB.Create('Les salaries', nil, -1); //PT17
  T1.LoadDetailDB('SALARIES', '', '', Q, FALSE, False);
  Ferme(Q);
  if T1.Detail.Count > 0 then
  begin
    T2 := T1.Detail[0];
    if DS.State in [dsInsert] then
    begin
      for num := 1 to 5 do
      begin
        Numero := InttoStr(num);
{$IFNDEF EAGLCLIENT}
        Mensuel := THDbEdit(GetControl('PCI_SALAIREMOIS' + Numero));
        Annuel := THDbEdit(GetControl('PCI_SALAIRANN' + Numero));
{$ELSE}
        Mensuel := THEdit(GetControl('PCI_SALAIREMOIS' + Numero));
        Annuel := THEdit(GetControl('PCI_SALAIRANN' + Numero));
{$ENDIF}
        if Mensuel <> nil then
        begin
          SetField('PCI_SALAIREMOIS' + Numero, T2.GetValue('PSA_SALAIREMOIS' + Numero)); // PT16
        end;
        if Annuel <> nil then
        begin
          SetField('PCI_SALAIRANN' + Numero, T2.GetValue('PSA_SALAIRANN' + Numero)); // PT16
        end;
      end;
      SetField('PCI_HORAIREMOIS', T2.GetValue('PSA_HORAIREMOIS')); //PT16
      SetField('PCI_HORHEBDO', T2.GetValue('PSA_HORHEBDO')); // PT16
      SetField('PCI_HORANNUEL', T2.GetValue('PSA_HORANNUEL')); //PT16
      SetField('PCI_TAUXHORAIRE', T2.GetValue('PSA_TAUXHORAIRE')); // PT16
      SetField('PCI_CONFIDENTIEL', T2.GetValue('PSA_CONFIDENTIEL'));
      SetField('PCI_DEBUTCONTRAT', T2.GetValue('PSA_DATEENTREE')); // PT12
      SetField('PCI_FINCONTRAT', T2.GetValue('PSA_DATESORTIE')); // PT12
      SetField('PCI_LIBELLEEMPLOI', T2.GetValue('PSA_LIBELLEEMPLOI')); // PT12
      SetField('PCI_FINTRAVAIL', T2.GetValue('PSA_DATESORTIE')); // PT13
      SetField('PCI_MOTIFSORTIE', T2.GetValue('PSA_MOTIFSORTIE')); // PT15
    end;
    TFFiche(ECran).DisabledMajCaption := TRUE; // PT11
    if T2 <> nil then Ecran.Caption := 'Contrats de travail Paie & GRH : ' + Salarie + ' ' + T2.GetValue('PSA_LIBELLE') +
      ' ' + T2.GetValue('PSA_PRENOM');
    UpDateCaption(TFFiche(ECran));
  end; // Fin si query non nulle
  FreeAndNil(T1);
// FIN PT16
  if ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE="' + Salarie + '" AND PSE_INTERMITTENT="X"') then //PT3
  begin
    SetControlVisible('GRPBXIS', True);
    Q := OpenSQL('SELECT PSE_ISCATEG FROM DEPORTSAL WHERE PSE_SALARIE="' + Salarie + '"', true);
    if (Q.Findfield('PSE_ISCATEG').AsString = 'OUV') or (Q.FindField('PSE_ISCATEG').AsString = 'TEC') then
    begin
      SetField('PCI_ISCACHETS', '-');
      SetField('PCI_ISHEURES', 'X');
      SetControlEnabled('PCI_ISCACHETS', False);
    end;
    Ferme(Q);
  end
  else SetControlVisible('GRPBXIS', False);
end;

procedure TOM_CONTRATTRAVAIL.OnNewRecord;
var
  numero, st: string;
  num: Integer;
{$IFNDEF EAGLCLIENT}
  Mensuel, Annuel: THDBEdit;
{$ELSE}
  Mensuel, Annuel: THEdit;
{$ENDIF}
  Q: TQuery;
begin
  inherited;
  SetField('PCI_TYPECONTRAT', 'CDI');
  SetControlProperty('PCI_NONPRECARITE', 'Value', '');
  for num := 1 to 5 do
  begin
    Numero := InttoStr(num);
{$IFNDEF EAGLCLIENT}
    Mensuel := THDbEdit(GetControl('PCI_SALAIREMOIS' + Numero));
{$ELSE}
    Mensuel := THEdit(GetControl('PCI_SALAIREMOIS' + Numero));
{$ENDIF}
    if Mensuel <> nil then
    begin
      Mensuel.Visible := TRUE;
      SetField('PCI_SALAIREMOIS' + Numero, 0); // PT16
    end;
{$IFNDEF EAGLCLIENT}
    Annuel := THDbEdit(GetControl('PCI_SALAIRANN' + Numero));
{$ELSE}
    Annuel := THEdit(GetControl('PCI_SALAIRANN' + Numero));
{$ENDIF}
    if Annuel <> nil then
    begin
      Annuel.Visible := TRUE;
      SetField('PCI_SALAIRANN' + Numero, 0); // PT16
    end;
  end;
  SetField('PCI_HORAIREMOIS', 0);
  SetField('PCI_HORHEBDO', 0);
  SetField('PCI_HORANNUEL', 0);
  SetField('PCI_TAUXHORAIRE', 0);
  // PT6 17/12/2002 PH V591 Toutes les dates sont initialisées à idate1900 ou 2099 au lieu de null
  SetField('PCI_DEBUTCONTRAT', IDate1900);
  SetField('PCI_FINCONTRAT', IDate1900);
  SetField('PCI_ESSAIDEBUT', IDate1900);
  SetField('PCI_ESSAIFIN', IDate1900);
  SetField('PCI_RNVESSAIDEB', IDate1900);
  SetField('PCI_RNVESSAIFIN', IDate1900);
  SetField('PCI_ISCACHETS', '-'); //PT3
  SetField('PCI_ISHEURES', '-');
  SetField('PCI_FINTRAVAIL', IDate1900); //PT13
  //DEBUT PT8
  Q := OpenSQL('SELECT PSA_ETABLISSEMENT FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"', True);
  if not Q.eof then SetField('PCI_ETABLISSEMENT', Q.FindField('PSA_ETABLISSEMENT').AsString);
  ferme(Q);
  //FIN PT8
  // DEB PT18
  if GetParamsocSecur('SO_PGRECUPCONTRAT', FALSE) then
  begin
    st := 'SELECT PCI_TYPECONTRAT,PCI_MOTIFCTINT,PCI_NONPRECARITE FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + Salarie + '" ORDER BY PCI_DEBUTCONTRAT DESC';
    Q := OpenSQL(St, True);
    if not Q.eof then
    begin
      st := Q.FindField('PCI_TYPECONTRAT').AsString;
      SetField('PCI_TYPECONTRAT', st);
      if st = 'CCD' then // Uniquement dans le cas contrat de type CDD
      begin
        SetField('PCI_MOTIFCTINT', Q.FindField('PCI_MOTIFCTINT').AsString);
        SetField('PCI_NONPRECARITE', Q.FindField('PCI_NONPRECARITE').AsString);
      end;
    end;
    ferme(Q);
  end;
  // FIN PT18
end;

procedure TOM_CONTRATTRAVAIL.OnUpdateRecord;
var
  QQ, Q: TQuery;
  iMax: Integer;
  Debut, DebutEssai, DebutRNEssai, FinEssai, FinRNEssai: Variant;
  MessEssai, MessEssai2, MessRNVEssai, St, StWhere: string;
  reponse: word;
  Fin: Variant; //PT14
begin
  inherited;

  LastError := 0;
  // PT10 09/05/2004 PH V_50 FQ 11244 : les champs DATE sont initialisés avec une valeur cohérente
  if GetField('PCI_DEBUTCONTRAT') < 10 then
    SetField('PCI_DEBUTCONTRAT', IDate1900);
  if GetField('PCI_FINCONTRAT') < 10 then
    SetField('PCI_FINCONTRAT', IDate1900);
  if GetField('PCI_ESSAIDEBUT') < 10 then
    SetField('PCI_ESSAIDEBUT', IDate1900);
  if GetField('PCI_ESSAIFIN') < 10 then
    SetField('PCI_ESSAIFIN', IDate1900);
  if GetField('PCI_RNVESSAIDEB') < 10 then
    SetField('PCI_RNVESSAIDEB', IDate1900);
  if GetField('PCI_RNVESSAIFIN') < 10 then
    SetField('PCI_RNVESSAIFIN', IDate1900);
  // FIN P10
  if GetField('PCI_TYPECONTRAT') = '' then
  begin
    LastError := 1;
    LastErrorMsg := 'Le type du contrat est obligatoire';
  end;
  if (GetField('PCI_DEBUTCONTRAT') < 10) then
  begin
    LastError := 1;
    LastErrorMsg := 'La date de début du contrat est obligatoire';
  end;
  //DEBUT PT7
  if GetField('PCI_FINCONTRAT') = IDate1900 then
    StWhere := 'AND ((PCI_DEBUTCONTRAT<="' + UsdateTime(GetField('PCI_DEBUTCONTRAT')) + '" AND PCI_FINCONTRAT="' + UsDateTime(IDate1900) + '")' +
      ' OR (PCI_DEBUTCONTRAT<="' + UsdateTime(GetField('PCI_DEBUTCONTRAT')) + '" AND PCI_FINCONTRAT>="' + UsDateTime(getField('PCI_DEBUTCONTRAT')) + '")' +
      ' OR (PCI_FINCONTRAT="' + UsDateTime(IDate1900) + '" OR PCI_FINCONTRAT>="' + UsDateTime(getField('PCI_DEBUTCONTRAT')) + '")' +
      ' OR (PCI_DEBUTCONTRAT>="' + UsDateTime(getField('PCI_DEBUTCONTRAT')) + '"))'
  else StWhere := 'AND ((PCI_DEBUTCONTRAT<="' + UsdateTime(GetField('PCI_DEBUTCONTRAT')) + '" AND PCI_FINCONTRAT="' + UsDateTime(IDate1900) + '")' +
    ' OR (PCI_DEBUTCONTRAT<="' + UsdateTime(GetField('PCI_DEBUTCONTRAT')) + '" AND PCI_FINCONTRAT>="' + UsDateTime(getField('PCI_DEBUTCONTRAT')) + '")' +
      ' OR (PCI_DEBUTCONTRAT<="' + UsdateTime(GetField('PCI_FINCONTRAT')) + '" AND PCI_FINCONTRAT="' + UsDateTime(IDate1900) + '")' +
      ' OR (PCI_DEBUTCONTRAT<="' + UsdateTime(GetField('PCI_FINCONTRAT')) + '" AND PCI_FINCONTRAT>="' + UsDateTime(getField('PCI_FINCONTRAT')) + '")' +
      ' OR (PCI_DEBUTCONTRAT >="' + UsdateTime(GetField('PCI_DEBUTCONTRAT')) + '" AND PCI_FINCONTRAT <="' + UsdateTime(GetField('PCI_FINCONTRAT')) + '" AND PCI_FINCONTRAT<>"' +
      UsDateTime(IDate1900) + '"))';
  Q := OpenSQL('SELECT PCI_TYPECONTRAT,PCI_DEBUTCONTRAT,PCI_FINCONTRAT FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + Salarie + '" ' +
    'AND PCI_ORDRE<>' + IntToStr(GetField('PCI_ORDRE')) + ' ' + StWhere, True);
  if not Q.eof then
  begin
    if Q.FindField('PCI_FINCONTRAT').AsDateTime = IDate1900 then
      St := RechDom('PGTYPECONTRAT', Q.FindField('PCI_TYPECONTRAT').AsString, False) + ' au ' + DateToSTr(Q.FindField('PCI_DEBUTCONTRAT').AsDateTime)
    else
      St := RechDom('PGTYPECONTRAT', Q.FindField('PCI_TYPECONTRAT').AsString, False) + ' du ' + DateToSTr(Q.FindField('PCI_DEBUTCONTRAT').AsDateTime) + ' au ' +
        DateToSTr(Q.FindField('PCI_FINCONTRAT').AsDateTime);
    LastError := 1;
    PGIBox('Vous ne pouvez saisir de contrat sur cette période #13#10 car il existe déja un ' + St, Ecran.Caption);
  end;
  Ferme(Q);
  if LastError = 1 then Exit;
  //FIN PT7
  // DEBUT PT 3
  Debut := GetField('PCI_DEBUTCONTRAT');
  DebutEssai := GetField('PCI_ESSAIDEBUT');
  FinEssai := GetField('PCI_ESSAIFIN');
  DebutRNEssai := GetField('PCI_RNVESSAIDEB');
  FinRNEssai := GetField('PCI_RNVESSAIFIN');
  MessEssai := '';
  MessRNVEssai := '';
  MessEssai2 := '';
  //PT14
  Fin := GetField('PCI_FINCONTRAT');
  if ((GetField('PCI_TYPECONTRAT') = 'CCD') or (GetField('PCI_TYPECONTRAT') = 'CDD'))
    and (fin < 10) then
  begin
    PGIBox('Contrat à durée déterminée : la date de fin n''est pas renseignée #13#10 N''oubliez pas de la renseigner ultérieurement ' + St, Ecran.Caption);
  end;
  //fin PT14
  if (FinEssai < 10) and (DebutEssai > 10) then MessEssai := '#13#10 - La date de début est renseignée mais pas la date de fin#13#10';
  if (FinRNEssai < 10) and (DebutRNEssai > 10) then MessRNVEssai := '#13#10 - La date de début est renseignée mais pas la date de fin#13#10';
  if (FinEssai > 10) and (DebutEssai < 10) then MessEssai := MessEssai + '#13#10 - La date de fin est renseignée mais pas la date de début#13#10';
  if (FinRNEssai > 10) and (DebutRNEssai < 10) then MessRNVEssai := MessRNVEssai + '#13#10 - La date de fin est renseignée mais pas la date de début#13#10';
  if (FinEssai < DebutEssai) and (FinEssai > 10) and (DebutEssai > 10) then
    MessEssai := MessEssai + '#13#10- La date de fin "' + DateToStr(GetField('PCI_ESSAIFIN')) + '" ne peut être inférieure à la date de début "' +
      DateToStr(GetField('PCI_ESSAIDEBUT')) + '"#13#10';
  if (FinRNEssai < DebutRNEssai) and (FinRNEssai > 10) and (DebutRNEssai > 10) then
    MessRNVEssai := MessRNVEssai + '#13#10- La date de fin "' + DateToStr(GetField('PCI_RNVESSAIFIN')) + '" ne peut être inférieure à la date de début "' +
      DateToStr(GetField('PCI_RNVESSAIDEB')) + '"#13#10';
  if (DebutRNEssai < FinEssai) and (DebutRNEssai > 10) and (FinEssai > 10) then
    MessRNVEssai := MessRNVEssai + '#13#10- La date de début "' + DateToStr(GetField('PCI_RNVESSAIDEB')) + '" ne peut être inférieure à la date de fin de la période précédente "' +
      DateToStr(GetField('PCI_ESSAIFIN')) + '"#13#10';
  if (FinEssai < 10) and (DebutRNEssai > 10) and (FinRNEssai > 10) then
  begin
    MessEssai := MessEssai + '#13#10- La date de fin de la période d''essai n''est pas renseignée#13#10';
    MessEssai2 := '#13#10(la période de renouvellement sera réinitialisée.)';
  end;
  if (Debut <> DebutEssai) and (DebutEssai > 10) and (Debut > 10) then
  begin
    if DebutEssai < Debut then
    begin
      LastErrorMsg := 'la date de début de la période d''essai ne peut pas être inférieure à la date de début du contrat';
      LastError := 1;
      Exit;
    end
    else
      MessEssai := MessEssai + '#13#10- la date de début "' + DateToStr(GetField('PCI_ESSAIDEBUT')) + '" est différente de la date de début du contrat "' +
        DateToStr(GetField('PCI_DEBUTCONTRAT')) + '"#13#10';
  end;
  if MessEssai <> '' then
  begin
    reponse := HShowMessage('1;Période d''essai ;Attention, ' + MessEssai + '#13#10 ' +
      '#13#10Voulez-vous modifier cette période avant la validation ?' + MessEssai2 + ';Q;YN;Y;N;;;', '', '');
    if reponse = 6 then
    begin
      if MessEssai2 <> '' then
      begin // 09/10/2003
        SetField('PCI_RNVESSAIDEB', IDate1900);
        SetField('PCI_RNVESSAIFIN', IDate1900);
      end;
      SetField('PCI_ESSAIDEBUT', IDate1900);
      SetField('PCI_ESSAIFIN', IDate1900);
      SetFocusControl('PCI_ESSAIDEBUT');
      LastError := 1;
    end;
  end;
  if MessRNVEssai <> '' then
  begin
    reponse := HShowMessage('1;Période de renouvellement d''essai ;Attention, ' + MessRNVEssai + '#13#10 ' +
      'Voulez-vous modifier cette période avant la validation ?;Q;YN;Y;N;;;', '', '');
    if Reponse = 6 then
    begin
      SetField('PCI_RNVESSAIDEB', IDate1900);
      SetField('PCI_RNVESSAIFIN', IDate1900);
      SetFocusControl('PCI_RNVESSAIDEB');
      LastError := 1;
    end;
  end;
  // FIN PT3
  if LastError = 0 then
  begin
    if DS.State in [dsInsert] then
    begin
      QQ := OpenSQL('SELECT MAX(PCI_ORDRE) FROM CONTRATTRAVAIL WHERE PCI_SALARIE="' + Salarie + '"', TRUE);
      if not QQ.EOF then imax := QQ.Fields[0].AsInteger + 1 else iMax := 1;
      Ferme(QQ);
      SetField('PCI_ORDRE', IMax);
      SetField('PCI_SALARIE', Salarie);
    end;
  end;
end;

procedure TOM_CONTRATTRAVAIL.AccesCheckBoxIS(Sender: TObject); //PT3
var
  NomSender: string;
begin
{$IFNDEF EAGLCLIENT}
  NomSender := TDBCheckBox(Sender).Name;
{$ELSE}
  NomSender := TCheckBox(Sender).Name;
{$ENDIF}
  if NomSender = 'PCI_ISHEURES' then
  begin
    if GetControlText('PCI_ISHEURES') = 'X' then SetControlEnabled('PCI_ISCACHETS', False)
    else SetControlEnabled('PCI_ISCACHETS', True);
  end;
  if NomSender = 'PCI_ISCACHETS' then
  begin
    if GetControlText('PCI_ISCACHETS') = 'X' then SetControlEnabled('PCI_ISHEURES', False)
    else SetControlEnabled('PCI_ISHEURES', True);
  end;
end;

initialization
  registerclasses([TOM_CONTRATTRAVAIL]);
end.

