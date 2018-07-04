{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 30/07/2001
Modifié le ... : 30/07/2001
Description .. : TOM concernant HISTOSAISRUB
Mots clefs ... : PAIE;ACOMPTE
*****************************************************************}
{    TOM concernant HISTOSAISRUB
En fait, la TOM ne va traiter que le cas de la saisie des acomptes
car on doit pouvoir consulter les acomptes à partir du multicritere
et les modifier par la fiche ACOMPTE
En ce qui concerne, la saisie par rubriques, il y a une unit specifique.
Les autres cas : Frais professionnels, Traitement d'un fichier fourni par un
logiciel externe ne permettent aucune saisie.
Voir S'il y a lieu de faire une édition pour consulter les lignes avant de les
intéger dans la paie.
Cette TOM concerne uniquement la saisie des acomptes.
*****************************************************************
PT1-1 18/03/2002 SB V571 Fiche de bug n°10009 Ajout fiche ecran saisie code etab
PT1-2 18/03/2002 SB V571 Design
PT2   24/12/2002 PH V591 Portage CWAS et améliorations diverses
PT3   17/07/2003 SB V_42 FQ 10677 Controle de vraisemblance, salarié sans type de règlement des acomptes
PT4   07/08/2003 PH V_421 FQ 10774 Accès modif et suppression des lignes provenant fichier import
============= Attention, la TOM sert à la gestion des acomptes
      et à accèder aux lignes provenant d'un fichier d'import ======================
PT5   03/09/2003 PH V_421 Elipsis liste salarié rend Nom+Prenom
PT6   30/01/2004 SB V_50  Utilisation du bouton supprimer de la fiche
PT7   19/08/2004 PH V_50  FQ 11612 Rajout des contrôles de cohérences
PT8   02/09/2004 PH V_50  FQ 10218 Rajout de la confidentialité dans le contrôle du salarié
PT9   01/10/2004 PH V_50  FQ 11612 Controle date entree et date de l'acompte
PT10  12/10/2004 PH V_50  FQ 11529 Controle montant acompte > 0
PT11  22/04/2005 PH V_60 Filtrage salarié confidentialité établissement
PT12  26/01/2007 FC V_80 Mise en place du filtage habilitation pour les lookuplist
                        pour les critères code salarié uniquement
PT13  05/08/2008 JPP V_80 FQ 15557 Message 'Salarié inconnu ou bien vous n''avez pas les droits pour accéder à ce salarié' aussi en 2-tiers
}


unit UTOMHISTOSAISRUB;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Spin, lookup,
{$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, DBCtrls, Fiche,
{$ELSE}
  eFiche,
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOM, UTOB, HTB97, EntPaie, PgOutils, P5Util, P5def,PgOutils2;

type
  TOM_HISTOSAISRUB = class(TOM)
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(stArgument: string); override;
    procedure OnLoadRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnNewRecord; override;
    procedure OnClose; override;
    procedure OnDeleteRecord; override;
  private
    Salarie, LeTitre, PaiAcp: string;
    Mvt {,prenom PT1-2}: string;
    TEtab: TOB; // ,Tob_ACP
    LeQuinz, DateDeb, DateFin: TDateTime;
    LBSUPMVT: thlabel;
    procedure AfficheNomSalarie;
    procedure SalClick(Sender: TObject);
    procedure SalExit(Sender: TObject);
    procedure DateExit(Sender: TObject); // PT7
//    procedure BeforeDeleteRecord(Sender: TObject);
    procedure ControleCoherenceChamp;
  end;

implementation
{ TOM_HISTOSAISRUB }

procedure TOM_HISTOSAISRUB.OnArgument(stArgument: string);
var
{$IFNDEF EAGLCLIENT}
  Sal, DD: THDBEdit;
{$ELSE}
  Sal, DD: THEdit;
{$ENDIF}
  st: string;
  Q: TQuery;
begin
  inherited;
  // PT2   24/12/2002 PH V591 Portage CWAS et améliorations diverses
  DateDeb := Idate1900;
  // PT4   07/08/2003 PH V_421 FQ 10774 Accès modif et suppression des lignes provenant fichier import
  if StArgument <> '' then
  begin // Correction erreur de convertion car modif des acomptes a perturbé le traitement ligne import
    DateDeb := StrToDate(ReadTokenSt(StArgument));
    DateFin := StrToDate(ReadTokenSt(StArgument));
  end else DateFin := Idate2099;
  // FIN PT4
  // FIN PT2
  {Bdel := Ttoolbarbutton97(getcontrol('MYDELETE'));      PT6
  if Bdel <> nil then Bdel.OnClick := BeforeDeleteRecord;   }
  LBSUPMVT := thlabel(getcontrol('LBSUPMVT'));
  if LBSUPMVT <> nil then exit;
  TEtab := Tob.create('Les Etabs', nil, -1);
  st := 'SELECT * FROM ETABCOMPL';
  Q := OpenSql(st, TRUE);
  TEtab.LoadDetailDB('ETABCOMPL', '', '', Q, FALSE);
  Ferme(Q);
  Mvt := ReadTokenSt(StArgument);
  //LeQuinz := Date;
  // PT2   24/12/2002 PH V591 Portage CWAS et gestion arrondi
  LeQuinz := Arrondi(int((Datedeb + datefin) / 2), 0); // PORTAGECWAS
  LeTitre := 'Saisie des acomptes Paie & GRH le ' + DateToStr(LeQuinz);
  Letitre := TraduireMemoire(LeTitre);
  Ecran.Caption := LeTitre;
  UpdateCaption(TFFiche(Ecran));
{$IFNDEF EAGLCLIENT}
  Sal := THDBEdit(GetControl('PSD_SALARIE'));
{$ELSE}
  Sal := THEdit(GetControl('PSD_SALARIE'));
{$ENDIF}
  if Sal <> nil then
  begin
    Sal.OnElipsisClick := SalClick;
    Sal.OnExit := SalExit;
  end;
{$IFNDEF EAGLCLIENT}
  DD := THDBEdit(GetControl('PSD_DATEDEBUT'));
{$ELSE}
  DD := THEdit(GetControl('PSD_DATEDEBUT'));
{$ENDIF}
  if DD <> nil then DD.OnExit := DateExit;

  SetControlProperty('LIBSAL', 'Caption', ''); //PT1-2
end;
{
procedure TOM_HISTOSAISRUB.BeforeDeleteRecord(Sender: TObject);
var
  Bdel: ttoolbarbutton97;
  Init: word;
begin
  Init := MrYes;
  if getfield('PSD_TOPREGLE') = 'X' then
    Init := HShowMessage('1;Suppression d''un acompte;Attention, cet acompte a été réglé. #13#10' +
      ' Etes-vous sûr de vouloir le supprimer ?;N;YN;Y;', '', '');
  if Init <> mrYes then exit;
  Bdel := ttoolbarbutton97(getcontrol('BDELETE'));
  if Bdel <> nil then Bdel.onclick(sender);
end;
}
procedure TOM_HISTOSAISRUB.OnChangeField(F: TField);
var
  nb: double;
  DateSaisie: TDateTime;
begin
  inherited;
  if LBSUPMVT <> nil then exit;
  if (F.FieldName = 'PSD_SALARIE') then
  begin
    AfficheNomSalarie;
  end;
  if (F.FieldName = 'PSD_MONTANT') then
  begin
    // PT2   24/12/2002 PH V591 Portage CWAS et le montant est un double !
    nb := GetField('PSD_MONTANT');
    if nb < 0 then
    begin
      PgiBox('Vous devez renseigner un montant positif!', 'Montant de l''acompte : ' + FloatToStr(nb) + '');
      // PT2   24/12/2002 PH V591 Portage CWAS et le montant est un double !
      SetField('PSD_MONTANT', nb * -1); // suppression de FloatToStr
    end;
  end;
  if (F.FieldName = 'PSD_DATEDEBUT') then
  begin
    //Date d'acompte comprise dans la periode d'acompte défini sur le mul
    if (DS.State in [dsInsert]) then
    begin
      DateSaisie := GetField('PSD_DATEDEBUT');
      // PT2   24/12/2002 PH V591 Portage CWAS et
      if (DateDeb > IDate1900) and (DateFin > Idate1900) then
        if (DateSaisie < DateDeb) or (DateSaisie > DateFin) then
        begin
          PGIBox('La date d''acompte doit être comprise entre le ' + DateToStr(DateDeb) + ' et le ' + DateToStr(DateFin), 'Acompte');
          SetField('PSD_DATEDEBUT', DateDeb); // PORTAGECWAS
          // FIN PT2
        end;
       if GetField ('PSD_SALARIE') <> '' then ControleCoherenceChamp;  
    end;
    // Date de fin = date de debut
    SetField('PSD_DATEFIN', GetField('PSD_DATEDEBUT'));
  end;
  //DEB PT1-1
  if (F.FieldName = 'PSD_ETABLISSEMENT') and (GetControlText('PSD_SALARIE') <> '') and (GetControlText('PSD_ETABLISSEMENT') <> '') then
    if Rechdom('PGSALARIEETAB', GetControlText('PSD_SALARIE'), False) <> GetControlText('PSD_ETABLISSEMENT') then
    begin
      SetField('PSD_SALARIE', '');
      SetControlText('LIBSAL', '');
    end;
  //FIN PT1-1
end;

procedure TOM_HISTOSAISRUB.OnLoadRecord;
var
  typeAlim, stAlim: string;
{$IFNDEF EAGLCLIENT}
  TA: Thdbedit;
{$ELSE}
  TA: Thedit;
{$ENDIF}
  Alimentation: thlabel;
begin
  inherited;
  if LBSUPMVT = nil then
  begin
    if DS.State in [dsInsert] then
    begin
      SetControlEnabled('PSD_SALARIE', TRUE);
      SetControlEnabled('PSD_DATEDEBUT', TRUE);
    end
    else
    begin
      SetControlEnabled('PSD_SALARIE', FALSE);
      SetControlEnabled('PSD_DATEDEBUT', FALSE);
      SetControlEnabled('PSD_ETABLISSEMENT', False); //PT1-1
    end;
    AfficheNomSalarie;
  end
  else
  begin // on est dans le cas gestion des mvts importés via fichier
{$IFNDEF EAGLCLIENT}
    TA := thDBEdit(getcontrol('PSD_TYPALIMPAIE'));
{$ELSE}
    TA := thEdit(getcontrol('PSD_TYPALIMPAIE'));
{$ENDIF}
    if TA = nil then exit;
    // PT4   07/08/2003 PH V_421 FQ 10774 Affichage nom du salarié = ergonomie
    AfficheNomSalarie;
    // PT2   24/12/2002 PH V591 Portage CWAS
    typeAlim := GetControlText('PSD_TYPALIMPAIE'); // TA.text;
    setcontrolenabled('PSD_BASE', pos('B', typeAlim) > 0);
    setcontrolenabled('PSD_TAUX', pos('T', typeAlim) > 0);
    setcontrolenabled('PSD_COEFF', pos('C', typeAlim) > 0);
    setcontrolenabled('PSD_MONTANT', pos('M', typeAlim) > 0);
  end;
  Alimentation := thlabel(getcontrol('ALIMENTATION')); // @@@
  stAlim := '';
  if Alimentation <> nil then
  begin
    if pos('B', typeAlim) > 0 then stAlim := 'BASE';
    if pos('T', typeAlim) > 0 then
      if stAlim = '' then
        stAlim := 'TAUX' else stAlim := stAlim + ' * TAUX';
    if pos('C', typeAlim) > 0 then
      if stAlim = '' then
        stAlim := 'COEFFICIENT' else stAlim := stAlim + ' * COEFFICIENT';
    if pos('M', typeAlim) > 0 then
      if stAlim = '' then
        stAlim := 'MONTANT' else stAlim := stAlim + ' * MONTANT';
    Alimentation.caption := stalim; // @@@
  end;

end;

procedure TOM_HISTOSAISRUB.AfficheNomSalarie;
var
  Q: TQuery;
  st, sal, nom: string; //PT1-2 Ajout prenom
  TE: TOB;
begin
  Sal := GetField('PSD_SALARIE');
  if (isnumeric(Sal) and (VH_PAIE.PgTypeNumSal = 'NUM')) then Sal := ColleZeroDevant(StrToInt(Sal), 10);
  st := 'SELECT PSA_LIBELLE,PSA_PRENOM,PSA_ETABLISSEMENT,PSA_PAIACOMPTE,PSA_TYPPAIACOMPT ' +
    'FROM SALARIES WHERE PSA_SALARIE="' + sal + '"';
  Q := OpenSQL(st, TRUE);
  // PT4   07/08/2003 PH V_421 FQ 10774 Accès modif et suppression des lignes provenant fichier import
  if LBSUPMVT <> nil then
  begin
    if not Q.EOF then TFFiche(Ecran).caption := 'Ligne import de ' + Q.findfield('PSA_LIBELLE').AsString + ' ' + Q.findfield('PSA_PRENOM').AsString
    else TFFiche(Ecran).caption := 'Ligne import salarié inconnu';
  end
    // FIN PT4
  else
  begin
    PaiAcp := '';
    if not Q.Eof then
    begin
      if Q.FindField('PSA_TYPPAIACOMPT').AsString = 'PER' then
        PaiAcp := Q.FindField('PSA_PAIACOMPTE').AsString
      else
      begin
        TE := TEtab.FindFirst(['ETB_ETABLISSEMENT'], [Q.FindField('PSA_ETABLISSEMENT').AsString], FALSE);
        if TE <> nil then PaiAcp := TE.GetValue('ETB_PAIACOMPTE');
      end;
      nom := RechDom('PGSALARIE', GetField('PSD_SALARIE'), FALSE);
      // prenom:=Q.findfield('PSA_PRENOM').AsString; PT1-2
      SetControlProperty('LIBSAL', 'Caption', nom {+' '+prenom PT1-2});
    end; // Fin si query non nulle
    TFFiche(Ecran).caption := 'Saisie des acomptes : ' + Sal + ' ' + nom {+' '+prenom PT1-2 };
  end;
  UpdateCaption(Ecran);
  Ferme(Q);
end;

procedure TOM_HISTOSAISRUB.OnNewRecord;
begin
  inherited;
  if LBSUPMVT <> nil then exit;
  SetField('PSD_ORIGINEMVT', 'ACP');

  SetField('PSD_DATEDEBUT', LeQuinz);
  SetField('PSD_DATEFIN', LeQuinz);
  if VH_Paie.PgRubAcompte <> '' then SetField('PSD_RUBRIQUE', VH_Paie.PgRubAcompte)
  else SetField('PSD_RUBRIQUE', 'ZZZZ');
  SetControlEnabled('PSD_SALARIE', TRUE);
  SetControlEnabled('PSD_DATEDEBUT', TRUE);
  SetField('PSD_AREPORTER', 'NON');
  SetFocusControl('PSD_SALARIE');
// DEB PT11
  PGPositionneEtabUser (THValComboBox (PGGetControl ('PSD_ETABLISSEMENT', Ecran)));
  if (PGRendEtabUser() <> '') then SetField ('PSD_ETABLISSEMENT' , PGRendEtabUser());
// FIN PT11
end;

procedure TOM_HISTOSAISRUB.OnUpdateRecord;
var
  st, temp,StEtab: string;
  Q: TQuery;
  Sortie,Entree: TDateTime;
begin
  inherited;
  if LBSUPMVT <> nil then exit;
  // Dans le cas d'acompte payé par cheque ou especes alors on le considere comme payé car pas de traitement de validation
  if (DS.State = dsinsert) then
  begin
    if PaiAcp = '' then { DEB PT3 }
    begin
      PGIBox('Ce salarié n''est affecté à aucun mode de règlement des acomptes.', 'Saisie impossible');
      LastError := 1;
      LastErrorMsg := '';
      Exit;
    end
    else { FIN PT3 }
      if (PaiAcp <> '008') then
    begin
      SetField('PSD_DATEPAIEMENT', GetField('PSD_DATEDEBUT'));
      SetField('PSD_TOPREGLE', 'X');
    end
    else
      SetField('PSD_DATEPAIEMENT', IDate1900); // PORTAGECWAS
    st := 'SELECT MAX(PSD_ORDRE)  AS ORDRE ' + //,PSD_SALARIE,PSD_DATEDEBUT,PSD_DATEFIN,PSD_RUBRIQUE '+
    'FROM HISTOSAISRUB ' +
      'WHERE PSD_ORIGINEMVT = "ACP" ' +
      'AND PSD_SALARIE="' + GetField('PSD_SALARIE') + '" ' +
      'AND PSD_DATEDEBUT="' + UsDateTime(GetField('PSD_DATEDEBUT')) + '" ' +
      'AND PSD_DATEFIN="' + UsDateTime(GetField('PSD_DATEDEBUT')) + '" ' +
      'AND PSD_RUBRIQUE="' + GetField('PSD_RUBRIQUE') + '"';
    Q := OpenSql(st, TRUE);
    if not Q.EOF then SetField('PSD_ORDRE', Q.FindField('ORDRE').AsInteger + 1)
    else SetField('PSD_ORDRE', 1);
    Ferme(Q);
  end;
  {if t = nil then SetField ('PSD_ORDRE', 1)
  else
   begin
   Setfield('PSD_ORDRE',t.getvalue('PSD_ORDRE')+1);
   t.putvalue('PSD_ORDRE',t.getvalue('PSD_ORDRE')+1);
   end;
  }
  // PORTAGECWAS
  if THEdit(GetControl('PSD_DATEDEBUT')) = nil then
  begin
    LastError := 1;
    exit;
  end;
  LastError := 0;
  Salarie := GetField('PSD_SALARIE');
  if (isnumeric(Salarie) and (VH_PAIE.PgTypeNumSal = 'NUM')) then Salarie := ColleZeroDevant(StrToInt(Salarie), 10);
  if Salarie = '' then
  begin
    LastError := 2;
    LastErrorMsg := 'Vous devez saisir un code salarié';
    SetFocusControl('PSD_SALARIE');
    exit;
  end;
  st := 'SELECT PSA_DATESORTIE,PSA_DATEENTREE,PSA_SUSPENSIONPAIE,PSA_ETABLISSEMENT,PSA_CONFIDENTIEL' +
    ' FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"';
// DEB PT11
  StEtab := PGRendEtabUser();
  if StEtab <> ''  then St := st + ' AND PSA_ETABLISSEMENT="'+StEtab+'"';
// FIN PT11  
  Q := OpenSql(st, TRUE);
  if not Q.EOF then
  begin
    Sortie := Q.FindField('PSA_DATESORTIE').AsDateTime;
    if Sortie > 10 then // Date de sortie non renseignee attention au NULL
    begin
      if GetField('PSD_DATEDEBUT') > Sortie then // PT7
      begin
        LastError := 1;
        LastErrorMsg := 'La date de saisie est supérieure à la date de sortie du salarié ?';
        SetFocusControl('PSD_SALARIE');
      end;
      if Q.FindField('PSA_SUSPENSIONPAIE').AsString = 'X' then
      begin
        LastError := 2;
        LastErrorMsg := 'La paie du salarié est suspendue, vous ne pouvez pas faire d''acompte';
        SetFocusControl('PSD_SALARIE');
      end;
    end;
    // DEB PT9
    Entree := Q.FindField('PSA_DATEENTREE').AsDateTime;
    if Entree >= GetField('PSD_DATEDEBUT') then  // PT7
    begin
        LastError := 3;
        LastErrorMsg := 'La date d''entrée est supérieure à la date de l''acompte';
        SetFocusControl('PSD_DATEDEBUT');
    end;
    // FIN PT9
    // DEB PT10
    if GetField ('PSD_MONTANT') <= 0 then
    begin
      LastError := 2;
      LastErrorMsg := 'Vous devez renseigner l''acompte avec un montant supérieur à 0';
    end;
    // FIN PT10
    if LastError = 0 then
    begin
      SetField('PSD_ETABLISSEMENT', Q.FindField('PSA_ETABLISSEMENT').AsString);
      temp := Q.FindField('PSA_CONFIDENTIEL').Asstring;
      SetField('PSD_CONFIDENTIEL', temp);
      if VH_PAIE.PGRubAcompte <> '' then SetField('PSD_RUBRIQUE', VH_PAIE.PGRubAcompte)
      else SetField('PSD_RUBRIQUE', 'ZZZZ');
    end;
  end // PORTAGE CWAS deplacement du end  de la ligne 336 à 345
  else // PT7
  begin
    LastError := 1;
    LastErrorMsg := 'Le salarié est inconnu';
    SetFocusControl('PSD_SALARIE');
  end;

  Ferme(Q);
end;

procedure TOM_HISTOSAISRUB.SalClick(Sender: TObject);
var
  swhere: string;
  LaDate: TDateTime;
begin
  swhere := '';
  LaDate := GetField('PSD_DATEDEBUT'); // PT7
  swhere := '((PSA_DATESORTIE >= "' + UsDateTime(LaDate) + '") OR (PSA_DATESORTIE is NULL) OR  (PSA_DATESORTIE="' + UsDateTime(iDate1900) + '"))';
  if GetField('PSD_ETABLISSEMENT') <> '' then //PT1-1 Ajout Cond.
    swhere := swhere + ' AND PSA_ETABLISSEMENT="' + GetField('PSD_ETABLISSEMENT') + '"'; //PT1-1
  // PT5   03/09/2003 PH V_421 Elipsis liste salarié rend Nom+Prenom
  swhere := RecupClauseHabilitationLookupList(swhere);  //PT12
  LookupList(THEdit(GetControl('PSD_SALARIE')), 'Les Salariés', 'SALARIES', 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', swhere, 'PSA_SALARIE', TRUE, -1);
end;

procedure TOM_HISTOSAISRUB.SalExit(Sender: TObject);
var
  Etab: string;
  edit: THEdit;
begin
  if THEdit(GetControl('PSD_SALARIE')) <> nil then
  begin
    edit := THEdit(Sender); //Ajoute 0 devant code salarie
    if edit <> nil then
      if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
        edit.text := AffectDefautCode(edit, 10);
    Salarie := GetControlText('PSD_SALARIE');
    SetControlProperty('LIBSAL', 'Caption', RechDom('PGSALARIE', Salarie, FALSE) {+' '+prenom });
    //DEB PT1-1
    Etab := RechDom('PGSALARIEETAB', GetControlText('PSD_SALARIE'), False);
    if (GetControlText('PSD_ETABLISSEMENT') <> Etab) and (GetControlText('PSD_SALARIE') <> '') then
      SetField('PSD_ETABLISSEMENT', Etab);
    //FIN PT1-1
  end;
  if Salarie = '' then
  begin
    SetControlText('LIBSAL', '');
    exit;
  end; //PT1-1
  ControleCoherenceChamp; // PT7
end;

procedure TOM_HISTOSAISRUB.OnClose;
begin
  inherited;
  if TEtab <> nil then
  begin
    TEtab.free;
    TEtab := nil;
  end;
end;

{ DEB PT6 }

procedure TOM_HISTOSAISRUB.OnDeleteRecord;
var
  Init: word;
begin
  inherited;
  Init := MrYes;
  if getfield('PSD_TOPREGLE') = 'X' then
    Init := HShowMessage('1;Suppression d''un acompte;Attention, cet acompte a été réglé. #13#10' +
      ' Etes-vous sûr de vouloir le supprimer ?;N;YN;Y;', '', '');
  if Init <> mrYes then LastError := 1;
end;
{ FIN PT6 }

// DEB PT7

procedure TOM_HISTOSAISRUB.ControleCoherenceChamp;
var
  Q: TQuery;
  LaDate: TDateTime;
  St: string;
  SqlConfStr: string; // PT13
begin
  st := 'SELECT PSA_SALARIE,PSA_SUSPENSIONPAIE,PSA_DATESORTIE FROM SALARIES WHERE PSA_SALARIE="' + Salarie + '"';
// DEB PT13
//{$IFNDEF EAGLCLIENT}
  SqlConfStr := SqlConf ('SALARIES');
  if SqlConfStr <> '' then
      st := st + ' AND '+ SqlConfStr;   // PT8
//{$ENDIF}
// FIN PT13
  Q := OpenSql(St, TRUE);
  if Q.EOF then
  begin // PT8 Contenu du message
    PGIBox('Salarié inconnu ou bien vous n''avez pas les droits pour accéder à ce salarié', Ecran.caption);
    SetFocusControl('PSD_SALARIE');
  end
  else
  begin
    if (Q.FindField('PSA_SUSPENSIONPAIE').AsString) = 'X' then
    begin
      PGIBox('Le salarié est en suspension de paie', Ecran.caption);
      SetFocusControl('PSD_SALARIE');
    end;
    LaDate := Q.FindField('PSA_DATESORTIE').AsDateTime;
    if (LaDate < GetField('PSD_DATEDEBUT')) and (LaDate > IDate1900) then
      PGIBox('La date de l''acompte est supérieure à la date de sortie du salarié ', Ecran.caption);
    if (LaDate > GetField('PSD_DATEDEBUT')) and (GetField('PSD_DATEDEBUT') > DebutDeMois(LaDate)) then
      PGIBox('Attention, vous allez faire un acompte au ' + DateToStr(GetField('PSD_DATEDEBUT')) + ' alors que le salarié sort le ' + DateToStr(LaDate), Ecran.caption);
  end;
  Ferme(Q);
end;

procedure TOM_HISTOSAISRUB.DateExit(Sender: TObject);
begin
  // ControleCoherenceChamp;
end;
// FIN PT7
initialization
  registerclasses([TOM_HISTOSAISRUB]);
end.

