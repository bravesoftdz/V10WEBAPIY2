{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 10/09/2001
Modifié le ... :   /  /
Description .. : Unit de gestion  du multi critère de la saisie par rubrique
Mots clefs ... : PAIE;PGDEPORTEE
*****************************************************************}
{
PT1 21/12/2001 PH V571 traitement de la sélection etablissement enlevé ==> géré par AGL
PT2 28/02/2002 PH V572 Multi selection sur les masques associés à l'utilisateur
PT3 13/05/2002 PH V575 Periode de la saisie correspond au mois en cours de l'exercice
PT4 21/05/2002 PH V575 Modif pour la saisie des primes Gestion de la liste et des controles
                       dans la même TOF
PT5 11/06/2002 PH V582 Modif pour la gestion des periodes la saisie des elts variables
PT6 18/06/2002 PH V582 Modif pour recupérer les masques de la saisie des primes <> de la saisie par rubriques
PT7 07/08/2002 PH V582 Modif gestion de la CheckBox liste des salariés avec ou saisie saisie des primes
PT8 03/07/2003 PH V_421 FQ 10762 Modif correction PT1 et si etab sélectionné alors suppression choix etablissement
PT9 23/09/2003 PH V_421 Passage de la clause where de la query en parametre pour ne traiter que les salariés séléctionnés
PT10 11/05/2004 PH V_50 FQ10997 Test si masque non vide
PT11 24/09/2004 PH V_50 ON force le traitement de tous les enregistrements de la liste
PT12 11/03/2005 JL V_60 Correction sélection salarié (déplacement GoToLeBookmark)
PT13 14/03/2005 JL V_60 Ajout gestion adjoint
PT14 03/08/2005 PH V_60 FQ 10997 Boucle sur les masques
PT15 04/08/2005 PH V_60 FQ 12468 Prise en compte des diff sélection des masques
PT16 14/09/2005 PH V_60 FQ 12506 Accès sélection établissement
PT17 23/11/2005 JL V_65 EManager : Ajout multi niveau pour responsable absence
PT18 29/05/2006 PH V_65 FQ 13204 Message alerte aucune sélection alors tous les salariés sont sélectionnés
PT19 10/01/2007 GGU V_80 FQ 13404 Saisie des bulletins complémentaires "par rubrique"
PT20 13/03/2007 GGU V_72 FQ 13972 Avertissement pour la saisie par rubrique sur une période de paie close
PT21 27/04/2007 JL V_720 Ajout gestion champs responsables absences et variables
PT22 27/04/2007 JL V_720 Modification du chargemetn pour saisie des primes ($IFDEF EPRIMES)
PT23 06/07/2007 FC V_72 FQ 14540 ergonomie
PT24 11/09/2007 GGU V_80 FQ 13972 Alerte lors d'une saisie par rubrique sur unr période de paie close
PT25 03/03/2008 FL  V_80 Emanager / Ajout de la consultation des primes
}
unit UTofPG_MulSaisRub;

interface
uses StdCtrls, Controls, Classes, Graphics, forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  HDB, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DBGrids, FE_Main, mul,
{$ELSE}
  MaineAGL, eMul,
{$ENDIF}
  HCtrls, HEnt1, EntPaie, Hqry, UTOF, UTOB, Lookup,
  Hmsgbox, AGLInit, ParamDat, P5Def, Windows, Messages,Ed_Tools,
  PgOutils, PGoutils2,UTobDebug;

type
  TOF_PGMULSAISRUB = class(TOF)
  private
    WW: THEdit;
    SaisieOk: Boolean;
    AvecContrl, CritSupp: string;
    DateDeb, DateFin: THEdit;
    NumMsq: string; // Numéro de masque associé à l'utilisateur
    Etab, Profil, Crit: string; // Champs definis par utilisateur par rapport au masque de saisie par rubrique
    LeType: string; // Identifiant si saisie par rubrique ou saisie des primes pour rajout dans le WHERE
    SQLGlobal : String ; // Chaine SQL globale correspondant à la concaténation des 3 SQL de sélection
    ATraiter: TCheckBox;
    AdjointVar, NumMasq: string; //PT13
    TypeSaisieRubrique : string; //PT19
    RequeteHistoRub : String;
{$IFDEF EPRIMES}
    LesRub: array[1..35] of string;
    LesMasq: array[1..5] of string; // codes des masques traités
    LesMontants: array[1..35] of double;
    LesDecal: array[1..5] of integer; // nbre de colonnes de décalage (décalge cumulé)
    NbCol: array[1..5] of integer; // nbre de colonnes ds chq masque
{$ENDIF EPRIMES}
    //    procedure AnneeChange(Sender: TObject);
    procedure ActiveWhere(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
    procedure DateDebExit(Sender: TObject);
    procedure DateFinExit(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    function RenseigneCrit(XX: string): string;
    function RenseigneProfil(XX: string): string;
    procedure OnExitSalarie(Sender: TObject);
    procedure ElipsisClickSal(Sender: TObject);
    procedure ClickATraiter(Sender: TObject);
    procedure AccesSalarie(Sender: TObject);
    procedure ClickHierarchie(Sender: TObject);
    procedure RespAbsElipsisClick(Sender : TObject); //PT21
    function ControlDatesCloturePaie(DateDebut, DateFin: TDateTime): boolean;
{$IFDEF EPRIMES}
    procedure Construiretob;
    procedure InitSaisieprimes;
    function RendSaisieRub(Salarie: string; Decalage, NbreCols: Integer; LeMasque,Tobprimes: TOB; var libelle: string): Boolean;
{$ENDIF EPRIMES}
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  end;
{$IFDEF EPRIMES}
  VAR TobSaiPrim,LesRem,TOB_Masque,T_RemMere : Tob;
{$ENDIF EPRIMES}

implementation

uses PgOutilseAgl, ParamSoc,P5Util;
//PT8 03/07/2003 PH V_421 FQ 10762 Modif correction PT1 et si etab sélectionné alors suppression choix etablissement

procedure TOF_PGMULSAISRUB.AccesSalarie(Sender: TObject);

begin
//  if Etab = '' then exit;
  LookupList(THEdit(Sender), 'Salarié', 'SALARIES', 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', SQLGlobal, 'PSA_SALARIE', True, -1);
end;

procedure TOF_PGMULSAISRUB.ActiveWhere(Sender: TObject);
var
  DDebut, DFin: TDateTime;
  StHierarchie,St: string;
begin
  if SaisieOk = FALSE then
  begin // La requete est faite de telle façon qu'elle ne rend rien
    WW.Text := ' PSA_ETABLISSEMENT="#@%" AND PSA_ETABLISSEMENT<>"#@%"';
    exit;
  end;
  DDebut := StrToDate(DateDeb.Text);
  DFin := StrToDate(DateFin.Text);
  if (WW <> nil) then
    WW.Text := '(PSA_SUSPENSIONPAIE <> "X") AND (PSA_DATEENTREE <="' + UsDateTime(DFin) + '") AND (((PSA_DATESORTIE >="' + UsDateTime(DDebut) +
      '")) OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "' + UsDateTime(iDate1900) + '"))';
  // PT4 21/05/2002 PH V575 Modif pour la saisie des primes Gestion de la liste et des controles
{$IFDEF EAGLCLIENT}
{$IFDEF EPRIMES}
  if (LeType = 'P') Or (LeType = 'C') then //PT25
  begin
    if not ConsultP then
    begin
      if AdjointVar = '' then //PT13
      begin
        //PT25 - Début
        If (LeType = 'P') Then
        Begin
            If TCheckBox(GetControl('CKHIERARCHIE')).State = CbChecked then
            StHierarchie :=  ' ((PSE_RESPONSVAR = "' + LeSalarie + '") OR '+
             '(PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL LEFT JOIN SERVICEORDRE ON PSE_CODESERVICE=PSO_CODESERVICE '+
             'WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+LeSalarie+'"))))'
            else StHierarchie :=  ' PSE_RESPONSVAR = "' + LeSalarie + '"';
        End
        Else
        Begin
        	// En consultation, on visualise les salariés dont on est responsables (absences) et ceux dont on valide les primes (VAR)
            If TCheckBox(GetControl('CKHIERARCHIE')).State = CbChecked then
            StHierarchie :=  ' ((PSE_RESPONSVAR="' + LeSalarie + '" OR PSE_RESPONSABS="' + LeSalarie + '") OR '+
             '(PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL LEFT JOIN SERVICEORDRE ON PSE_CODESERVICE=PSO_CODESERVICE '+
             'WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+LeSalarie+'" OR PGS_RESPONSABS="'+LeSalarie+'"))))'
            else StHierarchie :=  ' (PSE_RESPONSVAR = "' + LeSalarie + '" OR PSE_RESPONSABS="' + LeSalarie + '") ';
        End;
        //PT25 - Fin
        if WW.Text <> '' then WW.Text := WW.Text + ' AND' + StHierarchie
        else WW.Text := StHierarchie;
      end
      else
      begin
        If TCheckBox(GetControl('CKHIERARCHIE')).State = CbChecked then
            StHierarchie :=  ' PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL LEFT JOIN SERVICES ON PGS_CODESERVICE=PSE_CODESERVICE WHERE PGS_ADJOINTVAR="' + LeSalarie + '")'+
            ' OR '+
             '(PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL LEFT JOIN SERVICEORDRE ON PSE_CODESERVICE=PSO_CODESERVICE '+
             'WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_ADJOINTVAR="'+LeSalarie+'"))))'
        else StHierarchie :=  ' PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL LEFT JOIN SERVICES ON PGS_CODESERVICE=PSE_CODESERVICE WHERE PGS_ADJOINTVAR="' + LeSalarie + '")';
        if WW.Text <> '' then WW.Text := WW.Text + ' AND'+StHierarchie
        else WW.Text := StHierarchie;
      end;
    end
    else
    begin
         //DEBUT PT17
         If TCheckBox(GetControl('CKHIERARCHIE')).State = CbChecked then StHierarchie :=  ' ((PSE_RESPONSABS = "' + LeSalarie + '") OR '+
         '(PSA_SALARIE IN (SELECT PSE_SALARIE FROM DEPORTSAL LEFT JOIN SERVICEORDRE ON PSE_CODESERVICE=PSO_CODESERVICE '+
         'WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSABS="'+LeSalarie+'"))))'
         else StHierarchie :=  ' PSE_RESPONSABS = "' + LeSalarie + '"';
         if WW.Text <> '' then WW.Text := WW.Text + ' AND'+StHierarchie
         else WW.Text := StHierarchie;
         //FIN PT17
    end;
  end;
{$ENDIF}
{$ENDIF}
  if CritSupp <> '' then
  begin
    if WW.Text <> '' then WW.Text := WW.Text + ' AND ' + CritSupp
    else WW.Text := CritSupp;
  end;
  // FIN PT4 Gesttion du critère supp
  // PT4 21/05/2002 PH V575 Modif pour la saisie des primes Gestion de la liste et des controles
  if AvecContrl <> 'X' then
  begin
    if (ATraiter <> nil) and (LeType = 'P') then
    begin
      if ATraiter.Checked then
      begin // Saisie A Faire - liste des salaries sans saisie
        ATraiter.Caption := 'Liste des salariés sans saisie';
        st := '  (NOT EXISTS (SELECT PSP_SALARIE FROM HISTOSAISPRIM WHERE PSP_SALARIE=PSA_SALARIE AND ' +
          'PSP_DATEDEBUT >= "' + UsDatetime(DDebut) + '" AND PSP_DATEFIN <= "' + UsDatetime(DFIN) + '")' +
          ' OR (NOT exists (SELECT PSP_SALARIE FROM HISTOSAISPRIM WHERE PSP_SALARIE=PSA_SALARIE AND PSP_DATEDEBUT >= "' +
          UsDatetime(DDebut) + '" AND PSP_DATEFIN <= "' + UsDatetime(DFIN) +
          '" AND ((PSP_BASE<>0) OR (PSP_TAUX<>0) OR (PSP_COEFF<>0) OR (PSP_MONTANT<>0) ))))'
      end
      else
      begin
        ATraiter.Caption := 'Liste des salariés avec saisie';
        st := ' (EXISTS (SELECT PSP_SALARIE FROM HISTOSAISPRIM WHERE PSP_SALARIE=PSA_SALARIE AND ' +
          ' ((PSP_BASE<>0) OR (PSP_TAUX<>0) OR (PSP_COEFF<>0) OR (PSP_MONTANT<>0)) AND ' +
          'PSP_DATEDEBUT >= "' + UsDatetime(DDebut) + '" AND PSP_DATEFIN <= "' + UsDatetime(DFIN) + '"))';
      end;
      if WW.Text = '' then WW.Text := WW.Text + st
      else WW.Text := WW.Text + ' AND ' + st;
    end;
  end;
  if (Etab = '') and (Profil = '') and (Crit = '') and (ATraiter = nil) then exit;
{ DEB PT15
  if (Etab <> '') and (Profil = '') and (Crit = '') then
  begin
    // PT1 21/12/2001 PH V571 traitement de la sélection etablissement enlevé ==> géré par AGL
    //PT8 03/07/2003 PH V_421 FQ 10762 Modif correction PT1 et si etab sélectionné alors suppression choix etablissement
    WW.Text := WW.Text + ' AND PSA_ETABLISSEMENT="' + Etab + '"';
    exit;
  end;
  if (Crit <> '') and (Profil = '') then
  begin // Gestion du critere supplementaire
    // PT1 21/12/2001 PH V571 traitement de la sélection etablissement enlevé ==> géré par AGL
    //PT8 03/07/2003 PH V_421 FQ 10762 Modif correction PT1 et si etab sélectionné alors suppression choix etablissement
    if Etab <> '' then WW.Text := WW.Text + ' AND PSA_ETABLISSEMENT="' + Etab + '"';
    WW.Text := RenseigneCrit(WW.Text);
    exit;
  end;
  if (Profil <> '') and (Etab = '') and (Crit = '') then
  begin
    WW.Text := WW.Text + ' AND PSA_PROFIL="' + Profil + '" ';
    exit;
  end;
  if (Profil <> '') then
  begin
    // PT1 21/12/2001 PH V571 traitement de la sélection etablissement enlevé ==> géré par AGL
    //PT8 03/07/2003 PH V_421 FQ 10762 Modif correction PT1 et si etab sélectionné alors suppression choix etablissement
    if Etab <> '' then WW.Text := WW.Text + ' AND PSA_ETABLISSEMENT="' + Etab + '"';
    if Crit <> '' then WW.Text := RenseigneCrit(WW.Text);
    WW.Text := RenseigneProfil(WW.Text);
  end;
// FIN PT15
}

  if WW.Text <> '' then WW.Text := WW.Text + ' AND ';
  WW.Text := WW.Text + SQLGLOBAL;
end;


procedure TOF_PGMULSAISRUB.OnClose;
begin
  inherited;
{$IFDEF EPRIMES}
  If LesRem <> Nil then FreeAndNil(LesRem);
  If TOB_Masque <> Nil then FreeAndNil(TOB_Masque);
  If T_RemMere <> Nil then FreeAndNil(T_RemMere);
{$ENDIF EPRIMES}
end;

procedure TOF_PGMULSAISRUB.OnArgument(Arguments: string);
var
{$IFDEF EAGLCLIENT}
  Grille: THGrid;
{$ELSE}
  Grille: THDBGrid;
{$ENDIF}
  Q: TQuery;
  UserAppli : string;
  Num, i: Integer;
  Edit: THedit;
  DebPer, FinPer, ExerPerEncours, LeDroit, MasqueSaisie: string;
  OkOk: Boolean;
  AA, MM, JJ: WORD;
  Mois, St: string;
  LaDate: TDateTime;
  SQLEtab, SQLProfil, SQLCrit: string; // SQL correspondant aux critéres de sélection
  LeCrit : String;
  ChbxH        : TCheckBox; // PT@@
  Defaut : THEdit;
  MultiC : THMultiValComboBox;
  iChamp,r : Integer;
  StRem,rem : String;
begin
  inherited;
  AdjointVar := '';
{$IFDEF EPRIMES}
  if ExisteSQL('SELECT PGS_RESPONSVAR FROM SERVICES WHERE PGS_RESPONSVAR="' + V_PGI.UserSalarie + '"') then AdjointVar := '' //PT13
  else AdjointVar := V_PGI.UserSalarie;
  SetControlVisible('XX_VARIABLEDEB', FALSE);
  SetControlVisible('XX_VARIABLEFIN', FALSE);
  SetControlVisible('LBLPERIODE', FALSE);
  SetControlVisible('LBLA', FALSE);
    //DEBUT PT21
  If AdjointVar = '' then
  begin
       SetControlVisible('TPSE_RESPONSABS',True);
       SetControlVisible('PSE_RESPONSABS',True);
       SetControlVisible('TPSE_RESPONSVAR',True);
       SetControlVisible('PSE_RESPONSVAR',True);
       SetControlVisible('CKHIERARCHIE',True);
       Defaut:=ThEdit(getcontrol('PSE_RESPONSABS'));
       If Defaut<>nil then defaut.OnElipsisClick := RespAbsElipsisClick;
       MultiC := THMultiValComboBox(GetControl('PSE_RESPONSVAR'));
       SetControlEnabled('PSE_RESPONSABS',True);
       SetControlEnabled('PSE_RESPONSVAR',True);
       st := 'PGRESPONSVARIABLE';
       iChamp := TTToNum(St);
//       if iChamp > 0 then
//       begin
//                  V_PGI.DECombos[iChamp].Prefixe := 'PSI';
//         V_PGI.DECombos[iChamp].Where := '(PSI_TYPEINTERIM="EXT" OR PSI_TYPEINTERIM="SAL") AND PSI_INTERIMAIRE IN (SELECT PSE_RESPONSVAR FROM DEPORTSAL)&#@';
//       end;
//       existeSQL('select * from decombos where do_COMBO="PGRESPONSVARIABLE" AND DO_PREFIXE="PSA"')
       If V_PGI.DECombos[iChamp].Prefixe = 'PSA' then MultiC.plus :=  ' AND (PSA_SALARIE="'+V_PGI.UserSalarie+'" OR PSA_SALARIE IN (SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
       ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+V_PGI.UserSalarie+'")))'
       else MultiC.plus :=  ' AND (PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSVAR FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
       ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+V_PGI.UserSalarie+'")))';
  end;
  //FIN PT21
{$ENDIF}
  // PT4 21/05/2002 PH V575 Modif pour la saisie des primes Gestion de la liste et des controles
  st := Trim(Arguments);

  //PT19
  if pos('BCP',st) > 0 then
  begin
    TypeSaisieRubrique := 'BCP';  //"Saisie par rubrique" des Bulletin complémentaire
    // On Vire les 4 derniers caractères qui sont ;XXX avec XXX qui représente le
    // type de la saisie de rubrique (SRB ou BCP)
    LeType := '';
    AvecContrl := '';
    Ecran.Caption := 'Saisie groupée des bulletins complémentaires'; //PT23
    UpdateCaption(Ecran);
  end
  //PT19
  else if pos('SRB',st) > 0 then
  begin
    TypeSaisieRubrique := 'SRB'; //"Saisie par rubrique" de la Paie
    LeType := '';
    AvecContrl := '';
    Ecran.Caption := 'Saisie par rubrique';
    UpdateCaption(Ecran);
  end
  else
  begin
    TypeSaisieRubrique := 'PRM'; //Saisie des primes
    LeType := readtokenst(st);
    AvecContrl := readtokenst(st); // Recup du type de fonctionnement
  end;
  if AvecContrl = 'X' then SetControlVisible('CHBXATRAITER', FALSE);
  NumMasq := '';
{$IFDEF EAGLCLIENT}
  Grille := THGrid(GetControl('Fliste'));
{$ELSE}
  Grille := THDBGrid(GetControl('Fliste'));
{$ENDIF}
  if Grille <> nil then Grille.OnDblClick := GrilleDblClick;
  WW := THEdit(GetControl('XX_WHERE'));
  DateDeb := ThEdit(getcontrol('XX_VARIABLEDEB'));
  // PT7 18/06/2002 PH V582 Modif gestion de la CheckBox liste des salariés avec ou saisie saisie des primes
  ATraiter := TCheckBox(GetControl('CHBXATRAITER'));
  if ATraiter <> nil then ATraiter.OnClick := ClickATraiter;
  if DateDeb <> nil then DateDeb.OnExit := DateDebExit;
  DateFin := ThEdit(getcontrol('XX_VARIABLEFIN'));
  if DateFin <> nil then DateFin.OnExit := DateFinExit;
  // PT2 28/02/2002 PH V572 Multi selection sur les masques associés à l'utilisateur
  // PT6 18/06/2002 PH V582 Recup YDA_MASQUEVAR pour la saisie des primes
  SaisieOk := TRUE;
  Q := OpenSQL('select YDA_MASQUEPAIE,YDA_DROITUTIL,YDA_MASQUEVAR FROM DROITACCES WHERE YDA_PUTILISAT="' + V_PGI.User + '"', TRUE);
  if not q.eof then
  begin
    if (LeType = 'P') Or (LeType = 'C') then NumMasq := Q.FindField('YDA_MASQUEVAR').AsString //PT25
    else begin
    //PT19
     NumMasq := Q.FindField('YDA_MASQUEPAIE').AsString;
     if TypeSaisieRubrique = 'BCP' then         
        NumMasq := Q.FindField('YDA_MASQUEVAR').AsString  //Masque pour le bulletin complémentaire
     else
        NumMasq := Q.FindField('YDA_MASQUEPAIE').AsString; //Masque par défaut (pour la saisie par rubrique)
    //Fin PT19
    end;
    // FIN PT6
    LeDroit := Q.FindField('YDA_DROITUTIL').AsString;
  end
  else
  begin // PT10 11/05/2004 PH V_50 FQ10997 Test si masque non vide
    FERME(Q);
    PGIBOX('L''utilisateur ' + V_PGI.UserName + ' n''est pas autorisé à faire la saisie', Ecran.Caption);
    SaisieOk := FALSE;
    SetControlVisible('BOUVRIR', False);
  end;
  Ferme(Q);
  if NumMasq <> '' then
    if Pos(';', NumMasq) = 0 then NumMasq := NumMasq + ';';

  UserAppli := V_PGI.UserLogin;
  SaisieOk := TRUE;
  // PT2 28/02/2002 PH V572 Multi selection sur les masques associés à l'utilisateur
  if LeDroit <> 'X' then
  begin
    PGIBOX('L''utilisateur ' + V_PGI.UserName + ' n''est pas autorisé à faire la saisie', Ecran.Caption);
    SaisieOk := FALSE;
    SetControlVisible('BOUVRIR', False);
  end;
  i := 0;
  if NumMasq = '' then
  begin // utilisateur utilise tous les masques
    Q := OpenSQL('select PMR_ORDRE FROM MASQUESAISRUB ORDER BY PMR_ORDRE', TRUE);
    while not q.eof do
    begin
      NumMasq := NumMasq + Q.FindField('PMR_ORDRE').AsString + ';';
      i := i + 1;
      if i = 4 then break; // on autorise les 4 1ers masques uniquement pour dimensionner à 4*7=28 colonnes
      Q.NEXT ; // PT14
    end;
    Ferme(Q);
  end;
  NumMsq := NumMasq;
  MasqueSaisie := readtokenst(NumMsq);
  // DEB PT15
  if VH_Paie.PGCritSaisRub = 'OR1' then LeCrit := ' PSA_TRAVAILN1 ';
  if VH_Paie.PGCritSaisRub = 'OR2' then LeCrit := ' PSA_TRAVAILN2 ';
  if VH_Paie.PGCritSaisRub = 'OR3' then LeCrit := ' PSA_TRAVAILN3 ';
  if VH_Paie.PGCritSaisRub = 'OR4' then LeCrit := ' PSA_TRAVAILN4 ';
  if VH_Paie.PGCritSaisRub = 'STA' then LeCrit := ' PSA_CODESTAT';
  while MasqueSaisie <> '' do
  begin
    Q := OpenSQL('select PMR_ETABLISSEMENT,PMR_PROFIL,PMR_CRITEREORG FROM MASQUESAISRUB WHERE PMR_ORDRE="' + MasqueSaisie + '"', TRUE);
    if not q.eof then
    begin // recup des criteres du masque de saisie par rubrique
      if Q.Fields[0].AsString <> '' then Etab := Etab + Q.Fields[0].AsString +';'; // PT16
      if SQLEtab <> '' then SQLEtab := SQLEtab + ' OR ';
      if Q.Fields[0].AsString <> '' then SQLEtab := SQLEtab + ' PSA_ETABLISSEMENT = "'+ Q.Fields[0].AsString +'"';
      Profil := Profil + Q.Fields[1].AsString +';';
      if SQLProfil <> '' then SQLProfil := SQLProfil + ' OR ';
      if Q.Fields[1].AsString <> '' then SQLProfil := SQLProfil + ' PSA_PROFIL = "'+ Q.Fields[1].AsString +'"';
      Crit := Crit + Q.Fields[2].AsString +';';
      if SQLCrit <> '' then SQLCrit := SQLCrit + ' OR ';
      if Q.Fields[2].AsString <> '' then SQLCrit := SQLCrit + LeCrit + '= "'+ Q.Fields[2].AsString +'"';
    end;
    Ferme(Q);
    MasqueSaisie := readtokenst(NumMsq);
  end;
  if SQLEtab <> '' then SQLGlobal := '('+SqlEtab+')';
  if SQLProfil <> '' then
  begin
    if SQLGlobal <> '' then SQLGlobal := SQLGlobal + ' AND ';
    SQLGlobal := SQLGlobal + '('+SQLProfil+')';
  end;
  if SQLCrit <> '' then
  begin
    if SQLGlobal <> '' then SQLGlobal := SQLGlobal + ' AND ';
    SQLGlobal := SQLGlobal + '('+SQLCrit+')';
  end;

  // FIN PT15
  if (DateDeb <> nil) and (DateFin <> nil) then
  begin
    DateDeb.OnElipsisClick := DateElipsisclick;
    DateFin.OnElipsisClick := DateElipsisclick;
  end;
  // PT5 11/06/2002 PH V582 Modif pour la gestion de la saisie des elts variables
  if (LeType = 'P') Or (LeType = 'C') then OkOk := RendPeriodeSaisVar(ExerPerEncours, DebPer, FinPer) //PT25
  else OkOk := RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer);
  if OkOk then
  begin
    // PT3 13/05/2002 PH V575 Periode de la saisie correspond au mois en cours de l'exercice
    if DateDeb <> nil then DateDeb.text := DebPer;
    if DateFin <> nil then DateFin.text := FinPer;
    // PT19 Si on est en saisie par rubrique des bulletins COMPLEMENTAIRES
    // on initialise la date de fin égale à la date de début
    if (DateDeb <> nil) and (TypeSaisieRubrique = 'BCP') then
      DateFin.text := DebPer;
  end;
  // FIN PT5
  NumMsq := NumMasq; // recup de tous les masques utilisés
  for Num := 1 to 4 do
  begin
    VisibiliteChampSalarie(IntToStr(Num), GetControl('PSA_TRAVAILN' + IntToStr(Num)), GetControl('TPSA_TRAVAILN' + IntToStr(Num)));
    VisibiliteChampLibreSal(IntToStr(Num), GetControl('PSA_LIBREPCMB' + IntToStr(Num)), GetControl('TPSA_LIBREPCMB' + IntToStr(Num)));
  end;
  VisibiliteStat(GetControl('PSA_CODESTAT'), GetControl('TPSA_CODESTAT'));

  // PT4 21/05/2002 PH V575 Modif pour la saisie des primes Gestion de la liste et des controles
  Edit := THedit(getcontrol('PSA_SALARIE'));
  if Edit <> nil then
  begin
    Edit.OnExit := OnExitSalarie;
    if (LeType = 'P') Or (LeType = 'C') then Edit.OnElipsisClick := ElipsisClickSal //PT25
      //PT8 03/07/2003 PH V_421 FQ 10762 Filtrage pour ne voir que les salariés de l'etablissement
    else if Etab <> '' then Edit.OnElipsisClick := AccesSalarie;
  end;
  Edit := THedit(getcontrol('PSA_SALARIE_'));
  if Edit <> nil then
  begin
    Edit.OnExit := OnExitSalarie;
    if (LeType = 'P') Or (LeType = 'C') then Edit.OnElipsisClick := ElipsisClickSal //PT25
      //PT8 03/07/2003 PH V_421 FQ 10762 Filtrage pour ne voir que les salariés de l'etablissement
    else if Etab <> '' then Edit.OnElipsisClick := AccesSalarie;
  end;

  if (LeType = 'P') Or (LeType = 'C') then //PT25
  begin
    if (LeType = 'P') Then //PT25
        Ecran.Caption := 'Saisie ' + VH_Paie.PgLibSaisPrim + ' responsable : ' + RechDom('TTUTILISATEUR', V_PGI.User, FALSE)
    Else
    Begin
        Ecran.Caption := 'Consultation des primes du responsable : ' + RechDom('TTUTILISATEUR', V_PGI.User, FALSE);
        If ATraiter <> Nil Then ATraiter.Visible := False;
    End;
    UpdateCaption(Ecran);
    LaDate := StrToDate(getcontroltext('XX_VARIABLEFIN')); // Recup Date fin
    DecodeDate(LaDate, AA, MM, JJ);
    Mois := IntToStr(MM);
    if Length(Mois) = 1 then Mois := '0' + Mois;
    setControltext('LBLMOIS', GetControlText('LBLMOIS') + ' ' + RechDom('PGMOIS', Mois, FALSE) + ' ' + IntToStr(AA));
  end;
  //PT8 03/07/2003 PH V_421 FQ 10762 Filtrage sur etablissement unique donc pas visibilité des critères etablissement
  if Etab <> '' then
  begin
    SetControlVisible('PSA_ETABLISSEMENT', FALSE);
    SetControlVisible('PSA_ETABLISSEMENT_', FALSE);
    SetControlVisible('LBLETAB', FALSE);
    SetControlVisible('LBLETAB1', FALSE);
  end;
{ DEB PT@@@ }
{if getparamsocsecur('SO_IFDEFCEGID',False)=True then
  Begin
  ChbxH := TCheckBox (GetControl ('CKHIERARCHIE'));
  if Assigned(ChbxH) then
    Begin
    ChbxH.Visible := VH_PAIE.PGEcabHierarchie;  
    ChbxH.OnClick := ClickHierarchie;
    End;
  SetControlVisible('HIERARCHIE',VH_PAIE.PGEcabHierarchie);
  SetControlVisible('HHIERARCHIE',VH_PAIE.PGEcabHierarchie);
  End;
  }
{ FIN PT@@@ }
{$IFDEF EPRIMES}
  StRem := '';
  TOB_Masque := TOB.Create('Les caractéristiques du masque', nil, -1);
  st := 'SELECT * FROM MASQUESAISRUB ORDER BY PMR_ORDRE';
  Q := OpenSql(st, TRUE);
  TOB_Masque.LoadDetailDB('MASQUESAISRUB', '', '', Q, FALSE, FALSE);
  Ferme(Q);
  For i := 0 to TOB_Masque.Detail.Count - 1 do
  begin
       For r := 1 to 7 do
       begin
            Rem := TOB_Masque.Detail[i].Getvalue('PMR_COL'+IntToStr(r));
            If rem <> '' then StRem := StRem + ' "' + Rem + '",';
       end;
  end;
  if StRem <> '' then StRem := ' AND PRM_RUBRIQUE IN (' + Copy(StRem, 1, Length(StRem) - 1) + ')';
  LesRem := TOB.Create('Les Remunérations', nil, -1);
  Q := OpenSql('SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_TYPEBASE="00" OR PRM_TYPETAUX="00" OR PRM_TYPECOEFF="00" OR PRM_TYPEMONTANT="00"'+StRem +'', TRUE);
  LesRem.LoadDetailDB('REMUNERATION', '', '', Q, FALSE, FALSE);
  Ferme(Q);
{$ENDIF EPRIMES}
end;

procedure TOF_PGMULSAISRUB.GrilleDblClick(Sender: TObject);
var
  St, LeWhere, St1, Salarie: string;
  DDebut, DFin: TDateTime;
  BCherc: TToolbarButton97;
  i, rep: Integer;
  AncVal : Bool; //PT25
begin
  RequeteHistoRub := '';
  BCherc := TToolbarButton97(GetControl('BCHERCHE'));
  if BCherc = nil then exit;
  // PT10 11/05/2004 PH V_50 FQ10997 Test si masque non vide
  if not SaisieOk then exit;
  if (NumMsq = '') or (THQuery(Ecran.FindComponent('Q')) = nil) then exit;
  DDebut := StrToDate(DateDeb.Text);
  DFin := StrToDate(DateFin.Text);
  // DEB PT18
  if (TFMul(Ecran).FListe.nbSelected = 0) then
  begin
    If (LeType <> 'C') Then     //PT25
    Begin
        rep := PgiAsk ('Voulez-vous effectuer la saisie pour tous les salariés affichés ?',Ecran.caption);
        if rep <> mrYes then exit;
    End;
  end;
  // FIN PT18
  if (TFMul(Ecran).FListe.AllSelected) or (TFMul(Ecran).FListe.nbSelected = 0) then
  begin
{$IFDEF EAGLCLIENT}
    if (TFMul(Ecran).bSelectAll.Down) or (TFMul(Ecran).FListe.nbSelected = 0) then
      TFMul(Ecran).Fetchlestous;
{$ENDIF}
    TFmul(Ecran).Q.First;
    while not TFmul(Ecran).Q.EOF do
    begin
      Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
      St1 := St1 + ' PSA_SALARIE="' + Salarie + '" OR';
      RequeteHistoRub := RequeteHistoRub + ' "' + Salarie + '",';
      TFmul(Ecran).Q.Next;
    end;
    if St1 <> '' then St1 := ' (' + Copy(St1, 1, Length(st1) - 2) + ')';
    if RequeteHistoRub <> '' then RequeteHistoRub := ' PSP_SALARIE IN (' + Copy(RequeteHistoRub, 1, Length(RequeteHistoRub) - 1) + ')';
    CritSupp := St1;
  end
  else
  begin
    if (TFMul(Ecran).FListe.nbSelected > 0) then
    begin
//      j := TFMul(Ecran).FListe.NbSelected;
      for i := 0 to TFMul(Ecran).FListe.NbSelected - 1 do
      begin
        TFMul(Ecran).FListe.GotoLeBookmark(i); // PT12
{$IFDEF EAGLCLIENT}
        TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1);
{$ENDIF}
        Salarie := TFmul(Ecran).Q.FindField('PSA_SALARIE').asstring;
        St1 := St1 + ' PSA_SALARIE="' + Salarie + '" OR';
      end;
//      TFMul(Ecran).FListe.ClearSelected;
      if St1 <> '' then St1 := ' (' + Copy(St1, 1, Length(st1) - 2) + ')';
      CritSupp := St1;
      TFMul(Ecran).BCherche.Click;
    end
    else
    begin
      PgiBox('Vous devez sélectionner au moins un salarié', Ecran.Caption);
      Exit;
    end;
  end;
{$IFDEF EPRIMES}

  Construiretob;
  BCherc := TToolbarButton97(GetControl('BCHERCHE'));
  if BCherc = nil then exit;
  // PT10 11/05/2004 PH V_50 FQ10997 Test si masque non vide
  if not SaisieOk then exit;
//  if (NumMsq = '') or (THQuery(Ecran.FindComponent('Q')) = nil) then exit;
  DDebut := StrToDate(DateDeb.Text);
  DFin := StrToDate(DateFin.Text);
{$ENDIF EPRIMES}

  { Récupération de la Query pour traitement dans la fiche vierge }
{$IFDEF EAGLCLIENT}
  TheMulQ := TOB(Ecran.FindComponent('Q'));
{$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}

  NumMsq := StringReplace(NumMasq, ';', '+', [rfReplaceAll]); // substitution des ; par des + car c'est une seule zone
//  NumMsq[length(NumMsq)+1] := ';'; // pour que le readtokenst marche

  St := NumMsq + ';' +DateToStr(DDebut) + ';' + DateToStr(DFin);
  // PT9 23/09/2003 PH V_421 Passage de la clause where de la query en parametre pour ne traiter que les salariés séléctionnés
  LeWhere := RecupWhereCritere(TFMUL(ECran).Pages); // Recup de la clause where generee par le mul
  if LeWhere <> '' then st := st + ';' + LeWhere;
  // FIN PT9

  //PT24 Verification des dates de la saisie par rubrique /r paies cloturées
  if not ControlDatesCloturePaie(DDebut, DFin) then exit;

  if (LeType = 'P') Or (LeType = 'C') then       //PT25
  begin
  	// Sauvegarde du mode de consultation au cas où
    if (LeType = 'C') Then //PT25
    Begin
        AncVal   := ConsultP;
        ConsultP := True;
    End;
    AGLLanceFiche('PAY', 'SAISPRIM', '', '', St);
    if (LeType = 'C') Then //PT25
        ConsultP := AncVal;
  end
  else begin
   //PT19
   AGLLanceFiche('PAY', 'SAISRUB', '', '', St+';'+TypeSaisieRubrique);
//   AGLLanceFiche('PAY', 'SAISRUB', '', '', St);
   //Fin PT19
  end;
{$IFDEF EPRIMES}
TobSaiPrim.Free;
{$ENDIF EPRIMES}
  TheMulQ := nil;
  SetControlText('XX_WHERE', '');
  CritSupp := '';
  TFMul(Ecran).BCherche.Click;

  if (TFMul(Ecran).FListe.AllSelected = TRUE) then
  begin
    TFMul(Ecran).FListe.AllSelected := False;
    TFMul(Ecran).bSelectAll.Down := TFMul(Ecran).FListe.AllSelected;
  end;
{  else
    TFMul(Ecran).FListe.ClearSelected;}
end;

function TOF_PGMULSAISRUB.ControlDatesCloturePaie(DateDebut, DateFin: TDateTime): boolean;
var
  st, Cloture: string;
  Q: TQuery;
  AA, MD, MF, JJ: WORD;
  Indice, i: Integer;
begin
  result := FALSE;
  st := 'SELECT PEX_CLOTURE FROM EXERSOCIAL WHERE PEX_DATEDEBUT<="' + UsDateTime(DateDebut) + '" AND PEX_DATEFIN>="' + UsDateTime(DateFin) + '"' +
    ' ORDER BY PEX_DATEDEBUT DESC';
  Q := OpenSQL(st, TRUE);
  if not Q.EOF then
  begin
    Cloture := Q.FindField('PEX_CLOTURE').AsString;
    Ferme(Q);
  end else begin
    Ferme(Q);
    if PGIAsk('La période de saisie n''est pas incluse dans un exercice social.#13#10 Voulez vous continuer ?', 'Contrôle des dates de la saisie') = mrYes then
      result := TRUE;
    exit;
  end;
  // recuperation des mois
  DecodeDate(DateDebut, AA, MD, JJ);
  DecodeDate(DateFin, AA, MF, JJ);
  Indice := MD;
  if VH_Paie.PGDecalage = TRUE then
  begin
    if Indice = 1 then Indice := 2 else
    begin
      if Indice = 12 then Indice := 1 else
        Indice := Indice + 1;
    end;
  end;
  if MD = MF then // On a une saisie sur 1 mois
  begin
    if Cloture[Indice] = '-' then
      result := TRUE // OK mois non cloturé
    else if PGIAsk('La période de saisie est située sur un mois de paie clôturé.#13#10 Voulez vous continuer ?', 'Contrôle des dates de la saisie') = mrYes then
      result := TRUE;
  end
  else
  begin // on a une saisie sur plusieurs mois
    result := TRUE;
    for i := MD to MF do
    begin
      if VH_Paie.PGDecalage = TRUE then
      begin
        if i = 1 then Indice := 12 else
        begin
          if I = 12 then Indice := 1 else
            Indice := i + 1;
        end;
      end else Indice := i;
      if Cloture[Indice] = 'X' then
      begin
        if PGIAsk('La période de saisie est située sur un mois de paie clôturé.#13#10 Voulez vous continuer ?', 'Contrôle des dates de la saisie') = mrNo then
          Result := FALSE;
        break;
      end; // Si on trouve un seul mois cloturé alors erreur
    end;
  end;
end;

{
procedure TOF_PGMULSAISRUB.AnneeChange(Sender: TObject);
begin
  //ActiveWhere (Sender);
end;
}

procedure TOF_PGMULSAISRUB.DateDebExit(Sender: TObject);
begin
  if not IsValidDate(DateDeb.Text) then
  begin
    PGIBox('La date de début n''est pas valide', Ecran.Caption);
    DateDeb.SetFocus;
  end;
  //ActiveWhere (NIL);
end;

procedure TOF_PGMULSAISRUB.DateFinExit(Sender: TObject);
var
  Date1, Date2: TDateTime;
begin
  if not IsValidDate(DateFin.Text) then
  begin
    PGIBox('La date de fin n''est pas valide', Ecran.Caption);
    DateFin.SetFocus;
    exit;
  end;
  Date2 := StrToDate(DateFin.Text);
  Date1 := StrToDate(DateDeb.Text);
  if Date1 > Date2 then
  begin
    PGIBox('La date de début est supérieure à la date de fin', Ecran.Caption);
    DateDeb.SetFocus;
    DateFin.Text := '';
    exit;
  end;
  if (FINDEMOIS(Date1) <> FINDEMOIS(Date2)) then
  begin
    PGIBox('Les dates de début et de fin doivent être comprise dans le même mois', Ecran.Caption);
    DateFin.SetFocus;
    exit;
  end;
  //ActiveWhere (NIL);
end;

procedure TOF_PGMULSAISRUB.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;

function TOF_PGMULSAISRUB.RenseigneCrit(XX: string): string;
begin
  if VH_Paie.PGCritSaisRub = 'OR1' then XX := XX + ' AND PSA_TRAVAILN1="' + Crit + '"';
  if VH_Paie.PGCritSaisRub = 'OR2' then XX := XX + ' AND PSA_TRAVAILN2="' + Crit + '"';
  if VH_Paie.PGCritSaisRub = 'OR3' then XX := XX + ' AND PSA_TRAVAILN3="' + Crit + '"';
  if VH_Paie.PGCritSaisRub = 'OR4' then XX := XX + ' AND PSA_TRAVAILN4="' + Crit + '"';
  if VH_Paie.PGCritSaisRub = 'STA' then XX := XX + ' AND PSA_CODESTAT="' + Crit + '"';
  result := XX;
end;

function TOF_PGMULSAISRUB.RenseigneProfil(XX: string): string;
begin
  XX := XX + ' AND PSA_PROFIL="' + Profil + '" ';
  result := XX;
end;

procedure TOF_PGMULSAISRUB.OnLoad;
begin
  inherited;
  ActiveWhere(nil);

end;
// PT4 21/05/2002 PH V575 Modif pour la saisie des primes Gestion de la liste et des controles

procedure TOF_PGMULSAISRUB.OnExitSalarie(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGMULSAISRUB.ElipsisClickSal(Sender: TObject);
var StWhere: string;
begin
  If GetCheckBoxState('CKHIERARCHIE') = CbChecked then
  begin
    LookupList(THEdit(Sender), 'Salarié', 'SALARIES', 'PSA_SALARIE', 'PSA_LIBELLE,PSA_PRENOM', SQLGlobal, 'PSA_SALARIE', True, -1);
    exit
  end;
  if LeSalarie = '' then LeSalarie := '0000000002';
  if AdjointVar <> '' then //PT13
  begin
    StWhere := 'PSE_CODESERVICE IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_ADJOINTVAR="' + AdjointVar + '")';
    LookupList(TControl(Sender), 'Salarié', 'SALARIES LEFT JOIN DEPORTSAL ON PSA_SALARIE=PSE_SALARIE', 'DISTINCT PSE_SALARIE', 'PSA_LIBELLE', StWhere, 'PSE_SALARIE', True, -1);
  end
  else AfficheTabSalResp(Sender, LeSalarie, 'P');
end;
// FIN PT4 Nvelles Fonctions

procedure TOF_PGMULSAISRUB.ClickATraiter(Sender: TObject);
var
  BCherc: TToolbarButton97;
begin
  BCherc := TToolbarButton97(GetControl('BCHERCHE'));
  if BCherc = nil then exit;
  BCherc.Click;
  {
  if ATraiter <> NIL then
     begin
     if ATraiter.Checked then SetControlProperty ('CHBXATRAITER','Caption','Liste des salariés sans saisie')
        else SetControlProperty ('CHBXATRAITER','Caption','Liste des salariés ayant eu une saisie');
     end;
    }
end;

procedure TOF_PGMULSAISRUB.OnUpdate;
begin
  inherited;

  if AvecContrl = 'X' then
  begin
    //     GrilleDblClick(nil);
        //    PostMessage(TFMul(Ecran).Handle, WM_CLOSE, 0, 0);
  end;

end;

procedure TOF_PGMULSAISRUB.ClickHierarchie(Sender: TObject);
begin
SetcontrolEnabled('HIERARCHIE',(TCheckBox(Sender).Checked = True));
if TCheckBox(Sender).Checked = False then SetcontrolText('HIERARCHIE','');
end;

procedure TOF_PGMULSAISRUB.RespAbsElipsisClick(Sender : TObject); //PT19
var StWhere,StOrder : String;
begin
        StOrder := 'PSI_LIBELLE';
        {$IFDEF EPRIMES}
        if AdjointVar <> '' then StWhere := '(PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSABS FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
        ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_ADJOINTVAR="'+V_PGI.UserSalarie+'")))'
        else StWhere := '(PSI_INTERIMAIRE="'+V_PGI.UserSalarie+'" OR PSI_INTERIMAIRE IN (SELECT PGS_RESPONSABS FROM SERVICES LEFT JOIN SERVICEORDRE ON PGS_CODESERVICE=PSO_CODESERVICE'+
        ' WHERE PSO_SERVICESUP IN (SELECT PGS_CODESERVICE FROM SERVICES WHERE PGS_RESPONSVAR="'+V_PGI.UserSalarie+'")))';
        {$ENDIF}
        LookupList(THEdit(Sender),'Liste des responsables','INTERIMAIRES LEFT JOIN DEPORTSAL ON PSI_INTERIMAIRE=PSE_SALARIE','PSI_INTERIMAIRE','PSI_LIBELLE,PSI_PRENOM',StWhere,StOrder, True,-1);
end;

{$IFDEF EPRIMES}
procedure TOF_PGMULSAISRUB.Construiretob;
Var TobSaisieRubriques : Tob;
    TSal : Tob;
    DebutM1,FinM1,DDebut,DFin : TDateTime;
    LeWhere : String;
    aa,mm,jj : Word;
    SalAnn1: string;
    Mois : String;
    T_RemFille,LeMasque : Tob;
    st : String;
    Q : TQuery;
    PrimAffic: array[1..5] of Boolean; // colonnes elts de salaires affichables
    i,zz,j : Integer;
    LaRem : String;
    NbC,NbD,Ligne : Integer;
    
    LesTyp: array[1..35] of string;
    T1,T_Sal : Tob;
    LaRubrique : String;
    LibColRem: array[1..5] of string;
    PgNbSalLib,NbDecal,IndLign : Integer;


    QMul: ThQUERY; // Query recuperee du mul
    LesDecimales: array[1..35] of integer;
    LesAReporter: array[1..35] of string;
    okok : Boolean;
    Salarie,libelle,TypeCol : String;
    Ceg : Boolean;
    TS,TSRPrec,TSR : tOb;
    Montt: Double;

    Areprendre: Boolean;
    TobHistoPrim : Tob;
begin
  QMul := THQuery(Ecran.FindComponent('Q'));
  InitMoveProgressForm (NIL,'Chargement des données pour la saisie des primes',
                    'Veuillez patienter SVP ...', QMul.RecordCount,
                    False,True);
  TobSaiPrim := tob.Create('Saisie des primes',Nil,-1);
  TSal := nil;
  NumMsq := StringReplace(NumMsq, '+', ';', [rfReplaceAll]);
  DDebut := StrToDate(GetControlText('XX_VARIABLEDEB'));
  DFin := StrToDate(GetControlText('XX_VARIABLEFIN'));
  LeWhere := RecupWhereCritere(TFMUL(ECran).Pages);
  DebutM1 := DebutDeMois(PLUSMOIS(DDebut, -1));
  FinM1 := FINDEMOIS(PLUSMOIS(DFin, -1));
  DecodeDate(DDebut, AA, MM, JJ);
  // recuperation de la query du multicritere
  Mois := IntToStr(MM);
  if Length(Mois) = 1 then Mois := '0' + Mois;
  //if ConsultP then LaGrille.Options := LaGrille.Options-[goEditing];
  DecodeDate(DebutM1, AA, MM, JJ);
  InitTOB_Rem();
  // design des grille et definition des tailles des colonnes
{  for i := 3 to 37 do
  begin
    LaGrille.ColLengths[i] := 17;
  end;}
  If T_RemMere = Nil then
  begin
    T_RemMere := TOB.create('Les Rem des masques', nil, -1);
    for i := 1 to 4 do
    begin
      LesMasq[i] := ReadTokenSt(NumMsq); // codes des masques traités
      if LesMasq[i] = '' then break;
      LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[i]], FALSE);
      // Les décalages sont pour le traitement des colonnes suivantes donc indice +1 par rapport au masque en cours de traitement
      if LeMasque <> nil then
      begin
        NbC := LeMasque.getvalue('PMR_NBRECOL');
        LesDecal[i + 1] := LesDecal[i] + NbC;
        NbCol[i + 1] := NbCol[i + 1] + NbC;
        for zz := 1 to NbC do
        begin // recherche de la liste des remunerations contenues dans les masques
          LaRem := LeMasque.GetValue('PMR_COL' + IntToStr(zz));
          if LaRem <> '' then
          begin
            T_RemFille := T_RemMere.FindFirst(['CODEBASE'], [LaRem], FALSE);
            if T_RemFille = nil then
            begin
              T_RemFille := TOB.create('Une Rem', T_RemMere, -1);
              T_RemFille.AddChampSup('CODEBASE', False);
              T_RemFille.AddChampSup('LIBELBASE', False);
              T_RemFille.PutValue('CODEBASE', LaRem);
              T_RemFille.PutValue('LIBELBASE', LeMasque.GetValue('PMR_LIBCOL' + IntToStr(zz)));
            end;
          end;
        end;
      end;
    end;
  end;
  if LesMasq[1] = '' then
  begin
    PgiBox('Attention, aucun masque séléctionné', Ecran.caption);
    FiniMoveProgressForm();
    exit;
  end;
  // recup des remunerations ayant au moins un champ à saisir
  // pour connaitre les nbre de decimales à saisir dans la colonne en fonction du type de champ saisi

  TSal := TOB.Create('Les Salaries', nil, -1);
  St := 'SELECT PSA_SALARIE,PSA_CONFIDENTIEL,PSA_SALAIREMOIS1,PSA_SALAIREMOIS2,PSA_SALAIREMOIS3,PSA_SALAIREMOIS4,PSA_SALAIREMOIS5,PSA_BOOLLIBRE2,PSA_SALAIRANN1 FROM SALARIES ';
  st := st + 'LEFT JOIN DEPORTSAL ON PSE_SALARIE = PSA_SALARIE ';
  if LeWhere <> '' then st := st + LeWhere;
{$IFDEF EAGLCLIENT}
  if not ConsultP then
  begin
    if LeSalarie <> '' then st := st + ' AND PSE_RESPONSVAR = "' + LeSalarie + '"'
  end
  else
  begin
    if LeSalarie <> '' then st := st + ' AND PSE_RESPONSABS = "' + LeSalarie + '"';
  end;
{$ENDIF}
  St := St + ' ORDER BY PSA_SALARIE';
  Q := OpenSql(st, TRUE);
  TSal.LoadDetailDB('SALARIES', '', '', Q, FALSE, FALSE);
  Ferme(Q);

  PgNbSalLib := VH_Paie.PgNbSalLib;
  NbDecal := 0;
  for i := 1 to 5 do PrimAffic[i] := FALSE;

  for i := 1 to PgNbSalLib do
  begin
    if i = 1 then
      if VH_Paie.PgPrimAffichSal1 then PrimAffic[i] := TRUE;
    if i = 2 then
      if VH_Paie.PgPrimAffichSal2 then PrimAffic[i] := TRUE;
    if i = 3 then
      if VH_Paie.PgPrimAffichSal3 then PrimAffic[i] := TRUE;
    if i = 4 then
      if VH_Paie.PgPrimAffichSal4 then PrimAffic[i] := TRUE;
    if i = 5 then
      if VH_Paie.PgPrimAffichSal5 then PrimAffic[i] := TRUE;
  end;




  st := 'SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PGPRIMAFFICHSAL6"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then SalAnn1 := Q.FindField('SOC_DATA').AsString
  else SalAnn1 := '-';
  FERME(Q);
  if (SalAnn1 = 'X') then PrimAffic[5] := TRUE; // On force sur le 5eme élement
  for i := 1 to 5 do
  begin
    if PrimAffic[i] then NbDecal := NbDecal + 1;
  end;

  for i := 1 to PgNbSalLib do
  begin
    if i = 1 then
    begin
      if VH_Paie.PgPrimAffichSal1 then
        LibColRem[i] := VH_Paie.PgSalLib1
    end
    else if i = 2 then
    begin
      if VH_Paie.PgPrimAffichSal2 then
        LibColRem[i] := VH_Paie.PgSalLib2
    end
    else if i = 3 then
    begin
      if VH_Paie.PgPrimAffichSal3 then
        LibColRem[i] := VH_Paie.PgSalLib3
    end
    else if i = 4 then
    begin
      if VH_Paie.PgPrimAffichSal4 then
        LibColRem[i] := VH_Paie.PgSalLib4;
    end
    else if i = 5 then
    begin
      if VH_Paie.PgPrimAffichSal5 then
        LibColRem[i] := VH_Paie.PgSalLib5;
      if (SalAnn1 = 'X') then LibColRem[i] := 'Variable';
    end;
  end;

  for j := 1 to 4 do
  begin
    if LesMasq[j] = '' then break;
    LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[j]], FALSE);
    if LeMasque <> nil then
    begin
      for i := 1 to 7 do
      begin
        LaRubrique := LeMasque.GetValue('PMR_COL' + IntToStr(i));
        T1 := LesRem.FindFirst(['PRM_RUBRIQUE'], [LaRubrique], FALSE);
        if T1 <> nil then
        begin
          TypeCol := LeMasque.GetValue('PMR_TYPECOL' + IntToStr(i));
          if TypeCol = 'BAS' then NbD := T1.GetValue('PRM_DECBASE')
          else if TypeCol = 'COE' then NbD := T1.GetValue('PRM_DECCOEFF')
          else if TypeCol = 'TAU' then NbD := T1.GetValue('PRM_DECTAUX')
          else if TypeCol = 'MON' then NbD := T1.GetValue('PRM_DECMONTANT');
          LesRub[i + LesDecal[j]] := LeMasque.GetValue('PMR_COL' + IntToStr(i));
          LesTyp[i + LesDecal[j]] := LeMasque.GetValue('PMR_TYPECOL' + IntToStr(i));
          LesDecimales[i + LesDecal[j]] := NbD;
          LesAReporter[i + LesDecal[j]] := LeMasque.GetValue('PMR_REPORTCOL' + IntToStr(i)); // Stockage des colonnes à reporter
        end; // Fin si Rem connue
      end; // Fin Boucle sur nbre de colonnes gérées
    end; // si masque existe
  end; // boucle sur les masques
  // Boucle de chargement de la Grille

  Q := OpenSQL('SELECT * FROM HISTOSAISPRIM WHERE'+RequeteHistoRub+' AND PSP_DATEDEBUT ="' + USDateTime(DDebut)
  + '" AND PSP_DATEFIN ="' + USDateTime(DFin) + '" AND PSP_ORIGINEMVT="SRB" ORDER BY PSP_SALARIE,PSP_DATEDEBUT,PSP_DATEFIN,PSP_ORDRE',True);
  TobHistoPrim := TOB.Create('histo', nil, -1);
  TobHistoPrim.LoadDetailDB('HISTO', '', '', Q, FALSE, False);
  Ferme(Q);
  QMul.First;
  i := QMUL.RecordCount;
  Ceg := GetParamSocSecur('SO_IFDEFCEGID', FALSE);

  if VH_Paie.PgPrimMoisPrec then IndLign := 1
  else IndLign := 0;

  while not QMul.EOF do
  begin
    TSR := Tob.Create('Ligne prime',TobSaiPrim,-1);
    Salarie := QMul.FindField('PSA_SALARIE').AsString;
    TS := TSal.FindFirst(['PSA_SALARIE'],[Salarie],False);
    if VH_Paie.PgPrimMoisPrec then
    begin
      TSRPrec.AddChampSupValeur('SALARIE','-->Mois ' + IntToStr(MM) + '/' + IntToStr(AA));
      TSRPrec.AddChampSupValeur('NOM',QMul.FindField('PSA_LIBELLE').AsString + ' ' + QMul.FindField('PSA_PRENOM').AsString);
      TSRPrec.AddChampSupValeur('EMPLOI',RechDom('PGLIBEMPLOI', QMul.FindField('PSA_LIBELLEEMPLOI').AsString, FALSE));
    end;
    TSR.AddChampSupValeur('SALARIE',Salarie);
    TSR.AddChampSupValeur('NOM',QMul.FindField('PSA_LIBELLE').AsString + ' ' + QMul.FindField('PSA_PRENOM').AsString);
    TSR.AddChampSupValeur('EMPLOI',RechDom('PGLIBEMPLOI', QMul.FindField('PSA_LIBELLEEMPLOI').AsString, FALSE));
    j := 1;
    for i := 1 to 5 do
    begin
      if LibColRem[i] <> '' then
      begin
        if (i = 1) and Ceg then Montt := QMul.FindField('PSA_SALAIREMOIS1').AsFloat + QMul.FindField('PSA_SALAIREMOIS2').AsFloat + QMul.FindField('PSA_SALAIREMOIS4').AsFloat
        else Montt := QMul.FindField('PSA_SALAIREMOIS' + IntToStr(i)).AsFloat;
        TSR.AddChampSupValeur('CHAMP'+IntToStr(j),StrfMontant(Montt, 7, 2, '', TRUE));
        if (i = 5) and (SalAnn1 = 'X') then
           begin
             Montt := QMul.FindField('PSA_SALAIRANN1').AsFloat;
             if Montt = 1 then Montt := 0; // Pour forcer affichage
             TSR.PutValue('CHAMP'+IntToStr(j),StrfMontant(Montt, 7, 2, '', TRUE));
           end;
        j := j + 1;
      end;
    end;
    for i := 1 to 35 do LesMontants[i] := 0;

    for j := 1 to 4 do // boucle sur les masques
    begin
      if LesMasq[j] = '' then break;
      // REcupération des infos saisies le mois précédent
      if VH_Paie.PgPrimMoisPrec then
      begin
        RendSaisieRub(Salarie,LesDecal[j], NbCol[j + 1], LeMasque,TobHistoPrim, Libelle); // Recup de la saisie déjà effectuée le mois précédent
        LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[j]], FALSE);
        for i := 1 to LesDecal[j + 1] do
        begin
          TSRPrec.AddChampSupValeur('CHAMP'+IntToStr(i + NbDecal + LesDecal[j]),DoubleToCell(LesMontants[i + LesDecal[j]], LesDecimales[i + LesDecal[j]]));
        end;
      end;
      for i := 1 to 35 do LesMontants[i] := 0; // RAZ pour traiter le mois en cours
      // Fin traitement du mois précédent

      AReprendre := RendSaisieRub(Salarie, LesDecal[j], NbCol[j + 1], LeMasque, TobHistoPrim,Libelle); // Recup de la saisie déjà effectuée
//      if AReprendre = FALSE then RendSaisiePrec(Salarie, DateDebut, DateFin, LesDecal[j], NbCol[j + 1], LeMasque); // Recup de la saisie précédente
      for i := 1 to LesDecal[j + 1] do
      begin
//        if i + 2 + NbDecal + LesDecal[j] > LaGrille.ColCount then exit;
        if AReprendre = TRUE then TSR.AddChampSupValeur('CHAMP'+IntToStr(i + NbDecal + LesDecal[j]),DoubleToCell(LesMontants[i + LesDecal[j]], LesDecimales[i + LesDecal[j]]))
        else // Saisie à reprendre depuis la derniere saisie
        begin
          if LesAReporter[i + LesDecal[j]] = 'OUI' then
          begin // montant à reporter du mois précédent
            TSR.AddChampSupValeur('CHAMP'+IntToStr(i + NbDecal + LesDecal[j]),DoubleToCell(LesMontants[i + LesDecal[j]], LesDecimales[i + LesDecal[j]]));
          end
          else TSR.AddChampSupValeur('CHAMP'+IntToStr(i + NbDecal + LesDecal[j]),'');
        end;
      end; // fin boucle sur les colonnes
    end; // fin boucle sur les masques
    TSR.AddChampSupValeur('COMMENTAIRE',Libelle);
    QMul.NEXT;
    Ligne := Ligne + IndLign + 1;
    MoveCurProgressForm ('Salarié : '+Salarie);
  end; // si query non nulle

{  k := 1;
  for i := 3 + NbDecal to LaGrille.ColCount - 1 do
  begin // Boucle pour calculer à l'origine les valeurs en fct des taux ou coeff
    LeCode := CodeCal[k];
    if LeCode = '' then continue;
    CodeC := Readtokenst(LeCode);
    LeCode := Readtokenst(LeCode);
    if LeCode <> '0' then
    begin
      for zz := 1 to LaGrille.RowCount do
      begin
        Mt1 := Valeur(LaGrille.Cells[i, zz]);
        if Copy(LaGrille.Cells[0, zz], 1, 3) = '-->' then continue;
        TS := TSal.FindFirst(['PSA_SALARIE'], [LaGrille.Cells[0, zz]], FALSE);
        if TS <> nil then
        begin
          if StrToInt(LeCode) < 6 then Mt := TS.GetValue('PSA_SALAIREMOIS' + LeCode)
          else Mt := TS.GetValue('PSA_SALAIRANN1');
          if CodeC = '05' then
          begin
            //            if mt <> 1 then
            Mt1 := Mt * (Mt1 / 100);
            //             else mt := mt * (MT1 * 100) ;
          end
          else Mt1 := Mt * Mt1;
          Mt1 := ARRONDI(Mt1, 2);
          if mt <> 1 then LaGrille.cells[i, zz] := LaGrille.cells[i, zz] + '% soit ' + DoubleToCell(Mt1, 2)
          else LaGrille.cells[i, zz] := DoubleToCell(Mt1, 2);
          ;
        end;
      end;
    end;
    k := k + 1;
  end;                   }
  FiniMoveProgressForm();
  TobHistoPrim.Free;
end;

procedure TOF_PGMULSAISRUB.InitSaisieprimes;
begin
end;

function TOF_PGMULSAISRUB.RendSaisieRub(Salarie: string; Decalage, NbreCols: Integer; LeMasque,Tobprimes: TOB; var libelle: string): Boolean;
var
  Q: TQuery;
  St : string;
  i : Integer;
  Montant: Double;
  T : Tob;
begin
  result := FALSE;
  Libelle := ''; // #############
  For i :=  decalage to (Decalage + NbreCols) do
  begin

    T := Tobprimes.FindFirst(['PSP_SALARIE','PSP_ORDRE'],[Salarie,i],false);
    Montant := 0;
    If T <> NIL THEN
    begin
      if T.GetValue('PSP_TYPALIMPAIE') = 'BAS' then Montant := T.GetValue('PSP_BASE')
      else if T.GetValue('PSP_TYPALIMPAIE') = 'MON' then Montant := T.GetValue('PSP_MONTANT')
      else if T.GetValue('PSP_TYPALIMPAIE') = 'COE' then Montant := T.GetValue('PSP_COEFF')
      else if T.GetValue('PSP_TYPALIMPAIE') = 'TAU' then Montant := T.GetValue('PSP_TAUX');
      if (i > decalage) and (i <= decalage + Nbrecols) then
      begin
        if LesRub[i] = T.GetValue('PSP_RUBRIQUE') then
          LesMontants[i] := Montant;
      end;
      result := TRUE;
      Libelle := T.GetValue('PSP_LIBELLE');
    end;
  end;
  Ferme(Q);
end;
{$ENDIF EPRIMES}
initialization
  registerclasses([TOF_PGMULSAISRUB]);
end.

