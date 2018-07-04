{************UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion des exercices sociaux
Mots clefs ... : PAIE
*****************************************************************}
{
PT1   : 03/12/2001 SB V563 Chgmt de code : remplacement des .text par des
                           setfields
PT2   : 03/12/2001 JL V563 Fiches de bug 351 et 379 : contrôle des dates, et MAJ
                           des périodes de paie (sur le bouton suivant)
PT3   : 22/03/2002 JL V571 Fiche Bug N° 10019 : Exrcice actif cocher et
                           Etat=ouvert par défaut sur nouvel exercice
PT4   : 10/06/2002 PH V585 Gestion des périodes de la saisie des éléments
                           variables
PT5   : 19/09/2002 PH V585 Action uniquement sur les périodes de main dans le
                           menu outil
PT6   : 21/11/2002 PH V591 recuperation année par rapport à la date systeme pour
                           la premiere creation
PT7   : 30/12/2002 SB V591 FQ 10394 Ajout contrôle période saisie comprise dans
                           exercice défini
PT8   : 06/01/2003 PH V591 Fiche de bug n° 10404 : Message alerte controle
                           exercice avec décalage de paie
PT9   : 03/09/2003 PH V_42 compatible CWAS
PT10  : 06/08/2004 VG V_50 FQ N°11482
PT11  : 04/10/2004 PH V_50 Suppression rechargement tablette
PT12  : 08/10/2004 PH V_50 Mise à jour du jnal des evts
PT13  : 02/03/2005 VG V_60 En modification, interdire la modification de l'année
                           de référence - FQ N°11482
PT14  : 06/06/2005 PH V_60 FQ 11772 On ne peut pas créer ni supprimer en gestion
                           ds périodes
PT15  : 10/01/2006 VG V_65 Reprise PT13 - FQ N°11482
PT16 : 28/02/2006 EPI V_65 Remplacement Fiche liste par liste multicritère et fiche
                           FQ 12380 et 12738
PT17 : 30/03/2006 EPI V_65 FQ 13029 suppression dernier exercice à tort
														malgré le message d'avertissement
PT18 : 20/11/2007 FLO V_80 Affichage du dernier mois clos plutôt que la liste des mois clôturés
                            + contrôle que l'année de référence n'existe pas déjà en création
}

unit UTOMExerciceSocial;

interface
uses Windows, StdCtrls, Controls, Classes, sysutils,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB,Fiche,   // mise au point CWAS
{$ELSE}
  eFiche, uTob,            // mise au point CWAS
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, HTB97,  // mise au point CWAS
  ParamSoc, EntPaie, P5Def,
  PgOutils2;

type
  TOM_EXERSOCIAL = class(TOM)
  public
    procedure OnNewRecord; 										override;
    procedure OnArgument(stArgument: string); override;
    procedure OnUpdateRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnLoadRecord; override; //PT18

  private
    procedure OuverturePeriode(Sender: TObject);
    procedure ForceMAJPeriode(Sender: TObject); // PT2
  end;

implementation

Uses PGOutils;

{ Initialisation des valeurs par défaut
PH le 10/06/02 suppression des lignes de commentaires suite PT1
car + de commentaire que de codes ==> illisible
Rajout du PT4
}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_EXERSOCIAL.OnNewRecord;
var
  QAnnee: TQuery;
  Decalage: boolean;
  LaDateDuJour: TDateTime;
  SystemTime0: TSystemTime;
  aa, mm, jj: word;
  AnRef, Nbre: integer;
begin
  inherited;
  SetControlEnabled('PEX_ANNEEREFER', True); //PT13
  Decalage := VH_Paie.PGDecalage;
// PT6 21/11/2002 PH V591 recuperation année par rapport à la date systeme pour la premiere creation
  GetLocalTime(SystemTime0);
  LaDateDuJour := SystemTimeToDateTime(SystemTime0);
  decodedate(LaDateDuJour, aa, mm, jj);
  AnRef := aa;
  Nbre := 1;
  QAnnee := OpenSql('SELECT MAX(PEX_ANNEEREFER),MAX(PEX_EXERCICE) FROM EXERSOCIAL', True);
  if not QAnnee.EOF then
  begin
    if ((QAnnee.Fields[0].asstring) <> '') then
      AnRef := StrToInt(QAnnee.Fields[0].asstring) + 1;
    if ((QAnnee.Fields[1].asstring) <> '') then
      Nbre := StrToInt(QAnnee.Fields[1].asstring) + 1;
  end;
// FIN PT6
  Ferme(QAnnee);

  Setfield('PEX_ANNEEREFER', ColleZeroDevant(Anref, 4));
  Setfield('PEX_EXERCICE', ColleZeroDevant(Nbre, 3));
// PT9 03/09/2003 PH V_421 compatible CWAS
  if (AnRef <> 0) then
  begin
    if (Decalage = True) then
    begin
      Setfield('PEX_DATEDEBUT', (EncodeDate(Anref - 1, 12, 01)));
      Setfield('PEX_DATEFIN', (EncodeDate(Anref, 11, 30)));
    end;
    if (Decalage = False) then
    begin
      Setfield('PEX_DATEDEBUT', (EncodeDate(Anref, 01, 01)));
      Setfield('PEX_DATEFIN', (EncodeDate(Anref, 12, 31)));
    end;
    SetField('PEX_ETATEXERCICE', '001'); //PT3
    SetField('PEX_ACTIF', 'X');
  end;
  Setfield('PEX_DEBUTPERIODE', Getfield('PEX_DATEDEBUT'));
  Setfield('PEX_FINPERIODE', (FinDeMois(StrToDate(Getfield('PEX_DATEDEBUT')))));
// PT4 10/06/2002 PH V585 Gestion des périodes de la saisie des éléments variables
  Setfield('PEX_DEBUTSAISVAR', Getfield('PEX_DATEDEBUT'));
  Setfield('PEX_FINSAISVAR', (FinDeMois(StrToDate(Getfield('PEX_DATEDEBUT')))));
// FIN PT9
  SetField('PEX_CLOTURE', '------------');
  SetControlText('DERNIERMOISCLOS', ''); //PT18
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_EXERSOCIAL.OnArgument(stArgument: string);
var
  Btn, BtnV: TToolBarButton97;
  ChPaieDec: TCheckBox;
  St: string;
  // PT2
{$IFDEF EAGLCLIENT}
  DateDebutP, DateFinP: THEdit;
{$ELSE}
  DateDebutP, DateFinP: THDBEdit;
{$ENDIF}
begin
  inherited;
  SetControlEnabled('PEX_ANNEEREFER', False); //PT13
// PT4 10/06/2002 PH V585 Gestion des périodes de la saisie des éléments variables
  st := Trim(StArgument);
  if (St = 'V') then
  begin
    SetControlVisible('TPEX_DEBUTSAISVAR', TRUE);
    SetControlVisible('PEX_DEBUTSAISVAR', TRUE);
    SetControlVisible('TPEX_FINSAISVAR', TRUE);
    SetControlVisible('PEX_FINSAISVAR', TRUE);
    SetControlVisible('BPERIODEVAR', TRUE);
    BtnV := TToolBarButton97(GetControl('BPERIODEVAR'));
    if (BtnV <> nil) then BtnV.OnClick := OuverturePeriode;
  end
  else
  begin
    SetControlVisible('TPEX_DEBUTSAISVAR', FALSE);
    SetControlVisible('PEX_DEBUTSAISVAR', FALSE);
    SetControlVisible('TPEX_FINSAISVAR', FALSE);
    SetControlVisible('PEX_FINSAISVAR', FALSE);
    SetControlVisible('BPERIODEVAR', FALSE);
  end;
  if St = 'O' then
  begin // Cas où on ne modifie aucune zone on a acces au bouton periode
    SetControlEnabled('PEX_LIBELLE', False);
    SetControlEnabled('PEX_DATEDEBUT', False);
    SetControlEnabled('PEX_DATEFIN', False);
    SetControlEnabled('PEX_ETATEXERCICE', False);
    SetControlEnabled('PEX_ACTIF', False);
    SetControlEnabled('PEX_DEBUTPERIODE', False);
    SetControlEnabled('PEX_FINPERIODE', False);
    SetControlEnabled('PEX_DEBUTSAISVAR', False);
    SetControlEnabled('PEX_FINSAISVAR', False);
    SetControlEnabled('PEX_CLOTURE', False);
   // DEB PT14
    SetControlVisible('BDELETE', FALSE);
    SetControlVisible('BINSERT', FALSE);
  end;
  SetControlVisible('BIMPRIMER', FALSE);
   // FIN PT14
// FIN PT4
  ChPaieDec := TCheckBox(GetControl('CHPAIEDEC'));
  if ChPaieDec <> nil then ChPaieDec.Checked := VH_Paie.PGDecalage;

  Btn := TToolBarButton97(GetControl('BPERIODE'));
  if (Btn <> nil) then Btn.OnClick := OuverturePeriode;
  SetControlEnabled('PEX_CLOTURE', FALSE);

// PT2
{$IFNDEF EAGLCLIENT}
  DateDebutP := THDBEdit(GetControl('PEX_DEBUTPERIODE'));
  DateFinP := THDBEdit(GetControl('PEX_FINPERIODE'));
{$ELSE}
  DateDebutP := THEdit(GetControl('PEX_DEBUTPERIODE'));
  DateFinP := THEdit(GetControl('PEX_FINPERIODE'));
{$ENDIF}
 if DateDebutP <> nil then DateDebutP.OnChange := ForceMAJPeriode;
  if DateFinP <> nil then DateFinP.OnChange := ForceMAJPeriode;
// FIN PT2
end;

{ Incrémente d'un mois la periode en cours
PH Le 10/06/02 suppression des lignes en commentaire suite PT1
car + de lignes de commentaire que de code ==> illisible
et rajout du  PT4
}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_EXERSOCIAL.OuverturePeriode(Sender: TObject);
var
  DateFin, DebPer, FinPer: TDateTime;
{$IFNDEF EAGLCLIENT}
  Defaut: THDBValComboBox;
{$ELSE}
  Defaut: THValComboBox;
{$ENDIF}
  An, Mois, Jour: Word;
  St: string;
begin
  DebPer := idate1900;
  FinPer := idate1900;
{$IFDEF EAGLCLIENT}
    Defaut := THValComboBox(GetControl('PEX_ETATEXERCICE'));
{$ELSE}
    Defaut := THDBValComboBox(GetControl('PEX_ETATEXERCICE'));
{$ENDIF}
  // PT4 10/06/2002 PH V585 Gestion des périodes de la saisie des éléments variables
  if TToolBarButton97(Sender).name = 'BPERIODEVAR' then St := 'V' else st := '';

  if Defaut <> nil then Defaut.value := Defaut.Value;
  DateFin := GetField('PEX_DATEFIN');
  if (GetField('PEX_DEBUTPERIODE') > idate1900) and (GetField('PEX_FINPERIODE') > idate1900) then
  begin
    // PT4 10/06/2002 PH V585 Gestion des périodes de la saisie des éléments variables
    if St <> 'V' then
    begin
      DebPer := StrToDate(GetField('PEX_DEBUTPERIODE'));
      FinPer := StrToDate(GetField('PEX_FINPERIODE'));
    end
    else
    begin
      if (GetField('PEX_DEBUTSAISVAR') > idate1900) and (GetField('PEX_FINSAISVAR') > idate1900) then
      begin
        DebPer := StrToDate(GetField('PEX_DEBUTSAISVAR'));
        FinPer := StrToDate(GetField('PEX_FINSAISVAR'));
      end;
    end;
    // FIN PT4
    DecodeDate(FinPer, An, Mois, Jour);
    if (VH_Paie.PGDecalage = False) then
      if (Mois = 12) or (DateFin = FinPer) then exit;
    if (VH_Paie.PGDecalage = True) then
      if (Mois = 11) or (DateFin = FinPer) then exit;
    // PT4 10/06/2002 PH V585 Gestion des périodes de la saisie des éléments variables
    if St <> 'V' then
    begin
      //       SetField('PEX_DEBUTPERIODE',DateToStr(PlusMois(DebPer,1))); PORTAGECWAS
      //       SetField('PEX_FINPERIODE',DateToStr(FindeMois(PlusMois(FinPer,1))));
      SetField('PEX_DEBUTPERIODE', PlusMois(DebPer, 1));
      SetField('PEX_FINPERIODE', FindeMois(PlusMois(FinPer, 1)));
    end
    else
    begin
      //       SetField('PEX_DEBUTSAISVAR',DateToStr(PlusMois(DebPer,1))); PORTAGECWAS
      //       SetField('PEX_FINSAISVAR',DateToStr(FindeMois(PlusMois(FinPer,1))));
      SetField('PEX_DEBUTSAISVAR', PlusMois(DebPer, 1));
      SetField('PEX_FINSAISVAR', FindeMois(PlusMois(FinPer, 1)));
    end;
    // FIN PT4
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_EXERSOCIAL.OnDeleteRecord;
// PT16
var
 	QAnnee : Tquery;
	nb_exo : integer;
begin
  inherited;
// PT11 ChargementTablette(TFFicheListe(Ecran).TableName,'');
// PT10
// PT16 remplacement du test sur le nombre d'exercices
{  if (THGrid(TFFicheListe(Ecran).FListe).RowCount = 2) then
  begin
    PGIBox('Suppression impossible : Un exercice social est obligatoire',
      Ecran.Caption);
    DS.Insert;
  end;
}
//FIN PT10

// PT16
// recherche exercices en cours, il doit rester un exercice
   nb_exo := 0;
   QAnnee := OpenSQL('select PEX_EXERCICE FROM EXERSOCIAL ', TRUE);
   while not QAnnee.eof do
   begin
 	   nb_exo := nb_exo + 1;
     QAnnee.NEXT ;
   end;
   Ferme(QAnnee);
   if (nb_exo < 2) then
   begin
     LastError :=1;    // PT17
     PGIBox('Suppression impossible : Un exercice social est obligatoire',Ecran.Caption);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_EXERSOCIAL.OnUpdateRecord;
var
  Mess: string;
  DebExer: TDateTime;
  Jour, Mois, Annee: WORD;
  rep: Integer;
  TEvent: TStringList;
begin
  inherited;
//Rechargement des tablettes
  if (LastError = 0) and (Getfield('PEX_EXERCICE') <> '') and (Getfield('PEX_LIBELLE') <> '') then
// PT11 ChargementTablette(TFFicheListe(Ecran).TableName,'');

// PT2
    Mess := '';
  if ExisteSQL('SELECT PEX_EXERCICE FROM EXERSOCIAL WHERE PEX_EXERCICE<>"' + GetField('PEX_EXERCICE') + '" AND  ((PEX_DATEDEBUT<="' + UsDateTime(StrToDate(GetField('PEX_DATEDEBUT'))) + '" AND PEX_DATEFIN>="' + UsDateTime(StrToDate(GetField('PEX_DATEDEBUT'))) + '") OR (PEX_DATEDEBUT<="' + UsDateTime(StrToDate(GetField('PEX_DATEFIN'))) + '" AND PEX_DATEFIN>="' + UsDateTime(StrToDate(GetField('PEX_DATEFIN'))) + '"))') then
    Mess := '- Les dates de début et/ou de fin sont comprises dans un autre exercice.';
  if (StrToDate(GetField('PEX_DATEFIN'))) < (StrToDate(GetField('PEX_DATEDEBUT'))) then
    Mess := Mess + '#13#10 - La date de fin d''exercice ne peut être infèrieur à la date de début d''exercice.';
  if (StrToDate(GetField('PEX_FINPERIODE'))) < (StrToDate(GetField('PEX_DEBUTPERIODE'))) then
    Mess := Mess + '#13#10 - La date de fin de période ne peut être infèrieur à la date de début de période.';
  if Mess <> '' then
  begin
    LastError := 1;
    LastErrorMsg := (Mess);
  end;
// FIN PT2
// PT8 06/01/2003 PH V591 Fiche de bug n° 10404 : Message alerte controle exercice avec décalage de paie
  DebExer := StrToDate(GetControlText('PEX_DATEDEBUT'));
  if (DebExer <> Idate1900) then // exercice social existant
  begin
    DecodeDate(DebExer, Annee, Mois, Jour);
    if (Mois <> 12) then
    begin
      if (GetParamSoc('SO_PGDECALAGE')) then
      begin
        rep := PGIAsk('Attention, le dossier est en paie décalée et le mois de début #13#10' +
          'de l''exercice social n''est pas décembre ?', Ecran.caption);
        if rep <> mrYes then
        begin
          LastError := 2;
          SetFocusControl('PEX_DATEDEBUT');
          Exit;
        end;
      end;
    end;
    if (Mois <> 01) then
    begin
      if not (GetParamSoc('SO_PGDECALAGE')) then
      begin
        rep := PGIAsk('Attention, le dossier n''est pas en paie décalée et le mois de début #13#10' +
          'de l''exercice social n''est pas janvier ?', Ecran.caption);
        if rep <> mrYes then
        begin
          LastError := 2;
          SetFocusControl('PEX_DATEDEBUT');
          Exit;
        end;
      end;
    end;
  end;
  //PT18 - Début
  If (DS.State=dsInsert) And (ExisteSQL('SELECT 1 FROM EXERSOCIAL WHERE PEX_ANNEEREFER="'+GetField('PEX_ANNEEREFER')+'"')) Then
  Begin
      LastError := 2;
      LastErrorMsg := TraduireMemoire('Un autre exercice se base déjà sur cette année de référence.');
      SetFocusControl('PEX_ANNEEREFER');
      Exit;
  End;
  //PT18 - Fin
// FIN PT8
// DEB PT12
  if LastError <> 0 then
  begin
    TEvent := TStringList.Create;
    TEvent.Add('Modification exercice social ' + Getfield('PEX_LIBELLE'));
    CreeJnalEvt('003', '070', 'OK', nil, nil, TEvent);
    TEvent.free;
  end;
// FIN PT12
  SetControlEnabled('PEX_ANNEEREFER', False); //PT13
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_EXERSOCIAL.ForceMAJPeriode(Sender: TObject); // PT2
begin
   ForceUpdate;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE
*****************************************************************}

procedure TOM_EXERSOCIAL.OnChangeField(F: TField);
var
Datedeb, DateFin, DebPer, FinPer: TDateTime;
begin
inherited;
{DEB PT7 La période saisie doit être comprise dans l'exercice défini}
if (F.FieldName = 'PEX_DEBUTPERIODE') or (F.FieldName = 'PEX_FINPERIODE') then
   begin
   Datedeb := GetField('PEX_DATEDEBUT');
   DateFin := GetField('PEX_DATEFIN');
   DebPer := StrToDate(GetField('PEX_DEBUTPERIODE'));
   FinPer := StrToDate(GetField('PEX_FINPERIODE'));
   if ((DebPer < DateDeb) or (DebPer > DateFin)) and (DebPer > idate1900) then
      begin
      PgiBox('La date de début de période doit être comprise dans l''exercice social défini.', Ecran.Caption);
      SetField('PEX_DEBUTPERIODE', DateToStr(Datedeb));
      SetField('PEX_FINPERIODE', DateToStr(FindeMois(Datedeb)));
      end;
   if ((FinPer < DateDeb) or (FinPer > DateFin)) and (FinPer > idate1900) then
      begin
      PgiBox('La date de fin de période doit être comprise dans l''exercice social défini.', Ecran.Caption);
      SetField('PEX_FINPERIODE', DateToStr(DateFin));
      SetField('PEX_DEBUTPERIODE', DateToStr(DebutdeMois(DateFin)));
      end;
   end;
{FIN PT7 }
//PT15
if ((F.FieldName = 'PEX_ANNEEREFER') and (DS.State in [dsedit])) then
   SetControlEnabled('PEX_ANNEEREFER', False);
//FIN PT15

end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 20/11/2007 / PT18
Modifié le ... :   /  /
Description .. : Evènement appelé lors du chargement des données
Mots clefs ... :
*****************************************************************}
procedure TOM_EXERSOCIAL.OnLoadRecord; 
Begin
	SetControlText('DERNIERMOISCLOS', RendDernierMoisClos(GetField('PEX_CLOTURE'),VH_Paie.PGDecalage));
End;

initialization
  registerclasses([TOM_EXERSOCIAL]);

end.

