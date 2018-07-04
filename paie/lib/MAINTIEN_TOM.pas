{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 28/07/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : MAINTIEN (MAINTIEN)
Mots clefs ... : TOM;MAINTIEN
*****************************************************************}
{
PT1   : 28/12/2004 V_60   MF : Ajout de la possibilité de créer un
                               maintien, de le suprimer.
                               Recalcul du mt du maintien qd modif.
PT2   : 12/01/2005 V_60   MF : Récupération du Nom/prénom pour maj caption
PT3   : 24/02/2005 V_60   MF : Mise au point CWAS : corrige "conversion de type
                               variant incorrecte"  + sur garnatie message à tort
                               "saisir le type d'absence"
PT4   : 26/04/2005 V_602  MF  : FQ 12221 : CWAS : calcul nbre jours calendaires
PT5   : 01/06/2005 V_602  MF  : CWAS : nb jours calendaires et mt maintien
                                recalculés à tort qd on rentrait en modification
                                sur la fiche.
                                En modif de fiche il n'y a pas de calcul
                                automatique. Si besoin, il faut modifier
                                manuellement les champs à rectifier.
PT6   : 20/02/2006 V_650 MF   : FQ 12115 Les dates des reprises sont initialisées
                                à 31/12/1900. On ajoute un booléen pour repérer
                                les lignes de type historique (PMT_HISTOMAINT)
                                + Ajout d'un type de ligne (PMT_TYPEMAINTIEN)
                                pour différencier les lignes de maintien aux lignes
                                de garantie du salaire (remplace l'utilisation du
                                champ PMT_CARENCE à 9999)
PT7   : 20/02/2006 V_650 MF    : FQ 12121 : les lignes d'historique doivent être
                                 créées avant tout bulletin.
PT8   : 15/02/2007  V_702 MF  : correction on test le champ PMT_TYPEMAINTIEN et
                                non pas la valeur du contrôle (qui n'est pas sur la fiche)
                                + ferme (Q);
}
Unit MAINTIEN_TOM ;

Interface

Uses
//unused    StdCtrls,
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
//unused     FichList,
{$else}
     eFiche,
//unused     eFichList,
     UTob, //unused
{$ENDIF}
     EntPaie,
//unused     forms,
     sysutils,
//unused     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     PgOutils2,
     UTOM;
//unused     UTob ;

Type
  TOM_MAINTIEN = Class (TOM)
    private

    PsaLibelle, PsaPrenom, PmtSalarie      : string;
    State                                  : string;

    procedure ExitEdit(Sender: TObject);

    public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation

procedure TOM_MAINTIEN.OnNewRecord ;
begin
  Inherited ;
  if (State = 'ACTION=CREATION') then
    SetField('PMT_SALARIE','')
  else
   SetField('PMT_SALARIE',State);

  SetField('PMT_DATEDEBUTABS', IDate1900);
  SetField('PMT_DATEFINABS', IDate1900);
// d PT1
// d PT6
//  SetField('PMT_DATEDEBUT', StrToDate('01/01/1959'));
//  SetField('PMT_DATEFIN', StrToDate('01/01/1959'));

// En création les dates d'intégration ds la paie sont forcées à 31/12/1900
// afin de réserver la création à la saisie d'un historique
  SetField('PMT_DATEDEBUT', StrToDate('31/12/1900'));
  SetField('PMT_DATEFIN', StrToDate('31/12/1900'));
  SetField('PMT_HISTOMAINT', 'X');
  SetField('PMT_TYPEMAINTIEN','MAI');
// f PT6
// f PT1
  SetField('PMT_TYPECONGE','');
  SetField('PMT_LIBELLE','');
  SetField('PMT_BASEMAINTIEN',0);
  SetField('PMT_PCTMAINTIEN',0);
  SetField('PMT_CARENCE',0);
  SetField('PMT_NBJMAINTIEN',0);
  SetField('PMT_MTMAINTIEN',0);
  SetField('PMT_BASEABS',0);
  SetField('PMT_NBJABS',0);
  SetField('PMT_NBHEURES',0);
  SetField('PMT_RUBRIQUE','');

  State := 'ACTION=CREATION';
  SetcontrolVisible('TPMT_SALARIE', true);
  SetcontrolVisible('PMT_SALARIE', true);
  SetcontrolVisible('LSALARIE', true);

  SetFocusControl('PMT_SALARIE');
end ;

procedure TOM_MAINTIEN.OnDeleteRecord ;
begin
  Inherited ;
  if (GetField('PMT_DATEDEBUT') <> IDate1900) and
     (GetField('PMT_DATEFIN') <> IDate1900) then
  begin
// d PT1
// d PT6
//    if (GetField('PMT_DATEDEBUT') = strtodate('01/01/1959')) or
//       (GetField('PMT_DATEFIN') = strtodate('01/01/1959'))then
    if (GetField('PMT_HISTOMAINT') = 'X') then
// f PT6
    begin
      if not (PGIAsk('Il s''agit d''une ligne d''historique.#10#13'+
               ' (saisie manuelle).#10#13'+
               ' Confirmez-vous la suppression?','Maintien')= mrYes) then
      begin
        LastError := 1;
      end;
    end
    else
// f PT1
    begin
      if not (PGIAsk('Cette ligne de maintien a déjà fait l''objet d''une .#10#13'+
               ' intégration dans la paie.#10#13'+
               ' Confirmez-vous la suppression?','Maintien')= mrYes) then
      begin
        LastError := 1;
      end;
    end
  end;
end ;

procedure TOM_MAINTIEN.OnUpdateRecord ;
begin
  Inherited ;
  if (GetField('PMT_RUBRIQUE') = '') then
  begin
      SetFocusControl('PMT_RUBRIQUE');
      LastErrorMsg:='Renseigner la rubrique de maintien';
      LastError := 1;
  end;
// d  PT3
//  if (GetField('PMT_TYPECONGE') = '') then
// d PT6
//  if (GetField('PMT_CARENCE') <> 9999) and (GetField('PMT_TYPECONGE') = '') then
// ligne de type maintien
  if (GetField('PMT_TYPEMAINTIEN') = 'MAI') and (GetField('PMT_TYPECONGE') = '') then
// f PT6
// f PT3
  begin
      SetFocusControl('PMT_TYPECONGE');
      LastErrorMsg:='Saisir le type d''absence';
      LastError := 1;
  end;
  if (GetField('PMT_LIBELLE') = '') then
  begin
      SetFocusControl('PMT_LIBELLE');
      LastErrorMsg:='Saisir un libellé';
      LastError := 1;
  end;
  if (GetField('PMT_DATEFINABS') = IDate1900) then
  begin
      SetFocusControl('PMT_DATEFINABS');
      LastErrorMsg:='Il faut renseigner une date de fin d''absence';
      LastError := 1;
  end;
  if (GetField('PMT_DATEDEBUTABS') = IDate1900) then
  begin
      SetFocusControl('PMT_DATEDEBUTABS');
      LastErrorMsg:='Il faut renseigner une date de début d''absence';
      LastError := 1;
  end;
  if (GetField('PMT_SALARIE') = '') then
  begin
      SetFocusControl('PMT_SALARIE');
      LastErrorMsg:='Le matricule salarié est requis';
      LastError := 1;
  end;
end ;

procedure TOM_MAINTIEN.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;


procedure TOM_MAINTIEN.OnLoadRecord ;
var
  st                                    : string;
  Q                                     : TQuery;
begin
  Inherited ;
  if (State = 'ACTION=CREATION') and (GetField('PMT_SALARIE') <> '') then
  begin
    st := 'Maintien pour le salarié '+
          PmtSalarie + ' '+ PsaLibelle + ' '+ PsaPrenom;
    TraduireMemoire (st);
    Ecran.Caption := st;
    UpdateCaption(TFFiche(Ecran));
  end
  else
    if (State <> 'ACTION=CREATION') and (GetField('PMT_SALARIE') <> '')then
    begin
// d PT2
      st := 'SELECT PSA_LIBELLE, PSA_PRENOM FROM SALARIES WHERE '+
            'PSA_SALARIE="'+GetField('PMT_SALARIE')+'"';

      Q := OpenSql(st,True);
      if not Q.eof then
      begin
        PsaLibelle  := Q.FindField('PSA_LIBELLE').AsString;
        PsaPrenom := Q.FindField('PSA_PRENOM').AsString;
        st := 'Maintien pour le salarié '+
               GetField('PMT_SALARIE') + ' '+
               Q.FindField('PSA_LIBELLE').AsString +
               ' '+ Q.FindField('PSA_PRENOM').AsString;
        TraduireMemoire (st);
        Ecran.Caption := st;
        UpdateCaption(TFFiche(Ecran));
      end;
      ferme(Q);
// f PT2

    end;
//d PT1
// les dates d'intégration ds la paie ne sont pas affichées qd non renseignée ou
// = 01/01/1959 (historique)
  if (GetField('PMT_DATEDEBUT') = IDate1900) or
     (GetField('PMT_DATEFIN') = IDate1900) or
// d PT6
//     (GetField('PMT_DATEDEBUT') = strtodate('01/01/1959')) or
//     (GetField('PMT_DATEFIN') = strtodate('01/01/1959'))then
     (GetField('PMT_HISTOMAINT') = 'X') then
// f PT6
//f PT1
  begin
    SetControlVisible('PMT_DATEDEBUT',False);
    SetControlVisible('TPMT_DATEDEBUT',False);
    SetControlVisible('PMT_DATEFIN',False);
    SetControlVisible('TPMT_DATEFIN',False);
  end
  else
  begin
    SetControlVisible('PMT_DATEDEBUT',True);
    SetControlVisible('TPMT_DATEDEBUT',True);
    SetControlVisible('PMT_DATEFIN',True);
    SetControlVisible('TPMT_DATEFIN',True);
  end;

// d PT6
//  if (GetField('PMT_CARENCE') = 9999) then*
  if (GetField('PMT_TYPEMAINTIEN') = 'GAR') then
// f PT6
  // Garantie du Net ou du brut
  begin
    SetControlVisible('TPMT_TYPECONGE',False);
    SetControlVisible('PMT_TYPECONGE',False);
    SetControlVisible('TPMT_CARENCE',False);
    SetControlVisible('TPMT_NBJMAINTIEN',False);
    SetControlVisible('TPMT_PCTMAINTIEN',False);   // PT1
    SetControlVisible('TPMT_BASEMAINTIEN',False);
    SetControlVisible('PMT_BASEMAINTIEN',False);
    SetControlVisible('PMT_CARENCE',False);
    SetControlVisible('PMT_NBJMAINTIEN',False);
    SetControlVisible('PMT_PCTMAINTIEN',False);    // PT1
    SetControlVisible('LIB',False);    // PT1
    SetControlVisible('LJOUR1',False);
    SetControlVisible('LJOUR',False);
    SetControlVisible('MAINTIENNET',True);
    if (GetField('PMT_TYPECONGE') = 'NET') then
      SetControlProperty('MAINTIENNET','Caption','Garantie du net')
    else
      SetControlProperty('MAINTIENNET','Caption','Garantie du brut');
  end
  else
  // Maintien du salaire
  begin
    SetControlVisible('TPMT_TYPECONGE',True);
    SetControlVisible('PMT_TYPECONGE',True);
    SetControlVisible('TPMT_CARENCE',True);
    SetControlVisible('TPMT_NBJMAINTIEN',True);
    SetControlVisible('TPMT_PCTMAINTIEN',true);   // PT1
    SetControlVisible('TPMT_BASEMAINTIEN',True);
    SetControlVisible('PMT_BASEMAINTIEN',True);
    SetControlVisible('PMT_CARENCE',True);
    SetControlVisible('PMT_NBJMAINTIEN',True);
    SetControlVisible('PMT_PCTMAINTIEN',true);    // PT1
    SetControlVisible('LIB',true);    // PT1
    SetControlVisible('LJOUR1',True);
    SetControlVisible('LJOUR',True);
    SetControlVisible('MAINTIENNET',False);
  end;
end;

procedure TOM_MAINTIEN.OnChangeField ( F: TField ) ;
var
  st,Salarie                                    : string;
  Q                                             : TQuery;
begin
  Inherited ;

  if (State = 'ACTION=CREATION') and
//     (GetField('PMT_SALARIE') <> '') and
     (GetControlText('PMT_SALARIE') <> '') and
     ((F.FieldName = 'PMT_DATEDEBUTABS') or
     (F.FieldName = 'PMT_DATEFINABS')) then
  begin
    st := 'SELECT PMT_SALARIE FROM MAINTIEN WHERE '+
// PT4          'PMT_SALARIE = "'+ GetField('PMT_SALARIE')+'" AND '+
          'PMT_SALARIE = "'+ GetControlText('PMT_SALARIE')+'" AND '+
          '((PMT_DATEDEBUTABS <= "'+ UsDateTime(GetField('PMT_DATEDEBUTABS'))+ '" AND '+
          'PMT_DATEFINABS >= "'+ UsDateTime(GetField('PMT_DATEDEBUTABS'))+ '") OR '+
          '(PMT_DATEDEBUTABS >= "'+ UsDateTime(GetField('PMT_DATEDEBUTABS'))+ '" AND '+
          'PMT_DATEDEBUTABS <= "'+ UsDateTime(GetField('PMT_DATEFINABS'))+ '"))';


    Q := OpenSql(st,True);
    if not Q.eof then
    begin
      PGIBox ('Il existe déjà une ligne de maintien pour ces dates',
              'Dates de l''absence');
      SetFocusControl(F.FieldName );
    end;
    ferme (Q);
  end;

  if (F.FieldName = 'PMT_SALARIE') and
     (State = 'ACTION=CREATION') and
     (GetField('PMT_SALARIE') = '') then
    SetFocusControl(F.FieldName );
// d PT7
// les lignes d'historique doivent être saisies avant tout bulletin 
  if (F.FieldName = 'PMT_SALARIE') and
     (State = 'ACTION=CREATION') and
     (GetField('PMT_SALARIE') <> '') then
  begin
       st := 'SELECT PMT_SALARIE FROM MAINTIEN WHERE '+
             'PMT_SALARIE="'+GetField('PMT_SALARIE')+'"'+
             ' AND PMT_TYPEMAINTIEN = "MAI" AND PMT_HISTOMAINT = "-"';

       Q := OpenSql(st,True);
       if not Q.eof then
       begin
             PGIBox ('Attention, un calcul de maintien a déjà été effectué pour ce'+
                     ' salarié, vous devez supprimer le bulletin pour que l''historique'+
                     ' soit calculé correctement',
              'Reprise d''historique ');
       end ;
       ferme (Q); // PT8
  end;
// f PT7
  if (F.FieldName = 'PMT_SALARIE') and
     (GetField('PMT_SALARIE') <> '') then
  begin
    Salarie := Trim(GetField('PMT_SALARIE'));
    if (isnumeric(Salarie)) and (Salarie <> '') then
      if ((Length(salarie) = 10) and (salarie > '2147483647')) then
      begin
        PgiBox('Vous ne pouvez pas saisir un matricule superieur à 2147483647!', 'Matricule du salarié');
        SetField('PMT_SALARIE', '');
        SetFocusControl(F.FieldName )
      end;
  end;

  // nbre de jours de carence = 9999 : il s'agit d'une ligne de garantie
  if (F.FieldName = 'PMT_CARENCE') and
// d PT6
//     (GetField('PMT_CARENCE') = 9999) then
     (GetField('PMT_TYPEMAINTIEN') = 'GAR') then
// f PT6
  // Garantie  du salaire
  begin
    SetControlVisible('TPMT_TYPECONGE',False);
    SetControlVisible('PMT_TYPECONGE',False);
    SetControlVisible('TPMT_CARENCE',False);
    SetControlVisible('TPMT_NBJMAINTIEN',False);
    SetControlVisible('PMT_CARENCE',False);
    SetControlVisible('PMT_NBJMAINTIEN',False);
    SetControlVisible('LJOUR1',False);
    SetControlVisible('LJOUR',False);
    SetControlVisible('MAINTIENNET',True);
// PT4
    if (GetField('PMT_TYPECONGE') = 'NET') then
// PT6    if (GetControlText('PMT_TYPECONGE') = 'NET') then
      SetControlProperty('MAINTIENNET','Caption','Garantie du net')
    else
      SetControlProperty('MAINTIENNET','Caption','Garantie du brut');
  end;

// d PT1
  // date de début ou de fin de maintien
  if (State = 'ACTION=CREATION') and
     ((F.FieldName ='PMT_DATEDEBUTABS') or (F.FieldName ='PMT_DATEFINABS')) then
  begin
    // calcul du nbre de jours de maintien
(* PT4    SetField('PMT_NBJMAINTIEN', Getfield('PMT_DATEFINABS') -
                                Getfield('PMT_DATEDEBUTABS') + 1 -
                                GetField('PMT_CARENCE'))*)
    SetField('PMT_NBJMAINTIEN', StrToDate(GetControlText('PMT_DATEFINABS')) -
                                StrToDate(GetControlText('PMT_DATEDEBUTABS')) + 1 -
                                StrToInt(GetControlText('PMT_CARENCE')))
  end;

  // nbre jours de carence
  if (State = 'ACTION=CREATION') and    // PT5
     (F.FieldName = 'PMT_CARENCE') and
// PT4     (GetField('PMT_CARENCE') <> 9999) then
// d PT6
//     (StrToInt(GetControlText('PMT_CARENCE')) <> 9999) then

//PT8     (GetControlText('PMT_TYPEMAINTIEN') = 'MAI') then
     (GetField('PMT_TYPEMAINTIEN') = 'MAI') then
// f PT6
  // calcul du nbre de jours de maintien
  begin
(* PT4   Setfield('PMT_NBJMAINTIEN', Getfield('PMT_DATEFINABS') -
                                Getfield('PMT_DATEDEBUTABS') + 1 -
                                GetField('PMT_CARENCE'));*)
    Setfield('PMT_NBJMAINTIEN', StrToDate(GetControlText('PMT_DATEFINABS')) -
                                StrToDate(GetControlText('PMT_DATEDEBUTABS')) + 1 -
                                StrToInt(GetControlText('PMT_CARENCE')));
  end;

  // Nbre jours de carence, base du maintien, nbre jours de maintien, % maintien
  if (State = 'ACTION=CREATION') and    // PT5
     ((F.FieldName = 'PMT_CARENCE') or (F.FieldName = 'PMT_BASEMAINTIEN') or
      (F.FieldName = 'PMT_NBJMAINTIEN') or (F.FieldName = 'PMT_PCTMAINTIEN'))
      and
// PT4     (GetField('PMT_CARENCE') <> 9999) then
// d PT6
//   (StrToInt(GetControlText('PMT_CARENCE')) <> 9999) then
//PT8     (GetControlText('PMT_TYPEMAINTIEN') = 'MAI') then
     (GetField('PMT_TYPEMAINTIEN') = 'MAI') then
// f PT6
  // Calcul du montant du maintien
  // Mt du maintien = (base / Nbre jours calendaires absence) *
  //                          nbre de jours maintien *
  //                          % de maintien
  begin
// d PT3
    {$IFNDEF EAGLCLIENT}
    if ((GetField('PMT_BASEMAINTIEN ')<> '') and
        (Getfield('PMT_DATEFINABS') - Getfield('PMT_DATEDEBUTABS') + 1)*
         GetField('PMT_NBJMAINTIEN')* (GetField('PMT_PCTMAINTIEN')/100) <> 0) then
    {$ELSE}
    if ((GetField('PMT_BASEMAINTIEN ')<> 0) and
        (Getfield('PMT_DATEFINABS') - Getfield('PMT_DATEDEBUTABS') + 1)*
         GetField('PMT_NBJMAINTIEN')* (GetField('PMT_PCTMAINTIEN')/100) <> 0) then
    {$ENDIF}
// f PT3
    Setfield('PMT_MTMAINTIEN', Getfield('PMT_BASEMAINTIEN')/
                                 (Getfield('PMT_DATEFINABS') -
                                  Getfield('PMT_DATEDEBUTABS') + 1)*
                                  GetField('PMT_NBJMAINTIEN')*
                                  (GetField('PMT_PCTMAINTIEN')/100));
  end;
// f PT1
end ;

procedure TOM_MAINTIEN.OnArgument ( S: String ) ;
var
  st                                   : String;
// PT2
  Defaut                               : THEdit;
// PT2

begin
  Inherited ;
  PmtSalarie := '';
  PsaLibelle := '';
  PsaPrenom := '';

  st := Trim(S);

  State := ReadTokenSt(st);
  if (State <> 'ACTION=CREATION') then
  begin
    PmtSalarie := State;
    PsaLibelle := ReadTokenSt(st);
    PsaPrenom := ReadTokenSt(st);
  end;


  if (State <> 'ACTION=CREATION') then
  begin
//PT2??  st := 'Maintien pour le salarié '+
  st:=  PmtSalarie + ' '+ PsaLibelle + ' '+ PsaPrenom;
  TraduireMemoire (st);
  Ecran.Caption := st;
  UpdateCaption(TFFiche(Ecran));
  end;

  if (State = 'ACTION=CREATION') then
  begin
    SetcontrolVisible('TPMT_SALARIE', true);
    SetcontrolVisible('PMT_SALARIE', true);
    SetcontrolVisible('LSALARIE', true);
  end
  else
  begin
    SetcontrolVisible('TPMT_SALARIE', False);
    SetcontrolVisible('PMT_SALARIE', False);
    SetcontrolVisible('LSALARIE', False);
  end;

// d PT2
  Defaut:=ThEdit(getcontrol('PMT_SALARIE'));
  if Defaut<>nil then Defaut.OnExit:=ExitEdit;
// f PT2

end ;

procedure TOM_MAINTIEN.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_MAINTIEN.OnCancelRecord ;
begin
  Inherited ;
  OnLoadRecord;
end ;

procedure TOM_MAINTIEN.ExitEdit(Sender: TObject);
var
  edit                                          : thedit;
  st                                            : string;
  Q                                             : TQuery;

begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and
       (length(Edit.text)<11) and
       (isnumeric(edit.text)) then
      edit.text:=AffectDefautCode(edit,10);
// d PT2
    if (State = 'ACTION=CREATION') and
       (edit.text <> '') then
      begin
        st := 'SELECT PSA_LIBELLE, PSA_PRENOM FROM SALARIES WHERE '+
              'PSA_SALARIE="'+edit.text+'"';

        Q := OpenSql(st,True);
        if not Q.eof then
        begin
          PsaLibelle  := Q.FindField('PSA_LIBELLE').AsString;
          PsaPrenom := Q.FindField('PSA_PRENOM').AsString;

          st := 'Maintien pour le salarié '+
                edit.text + ' '+ Q.FindField('PSA_LIBELLE').AsString +
                ' '+ Q.FindField('PSA_PRENOM').AsString;
          TraduireMemoire (st);
          Ecran.Caption := st;
          UpdateCaption(TFFiche(Ecran));
        end
        else
        begin
          st := 'Maintien :';
          TraduireMemoire (st);
          Ecran.Caption := st;
          UpdateCaption(TFFiche(Ecran));
          PGIBox('Salarié inconnu ', Ecran.caption);
          SetField('PMT_SALARIE','');
          SetFocusControl('PMT_SALARIE');
        end;

        ferme(Q);

      end;
// f PT2

end;  { fin ExitEdit}

Initialization
  registerclasses ( [ TOM_MAINTIEN ] ) ;
end.

