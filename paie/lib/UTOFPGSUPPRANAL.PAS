{***********UNITE*************************************************
Auteur  ...... : FC
Créé le ...... : 29/05/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : UTOFPGSUPPRANAL ()
Mots clefs ... : TOF;UTOFPGSUPPRANAL
*****************************************************************}
Unit UTOFPGSUPPRANAL ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
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
     UTOF,
     HTB97,
     Ed_Tools,
     Menus,
     Ventil,
     EntPaie,
     PGOutils2,
     P5Def;

Type
  TOF_UTOFPGSUPPRANAL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
  private
    LeCompte : String;
    TVentilSal : TOB;
    procedure BDeleteClick(Sender: TObject);
    procedure SuppressionVentil;
    procedure MenuVentilClick(Sender: TObject);
    procedure RecupVentilSalarie;
    procedure DuplicationVentil;
    procedure ExitEdit(Sender: TObject);
  end ;

Implementation


procedure TOF_UTOFPGSUPPRANAL.OnArgument (S : String ) ;
var
  BDelete : TToolBarButton97;
  MenuVentil : TPopupMenu;
  Defaut : ThEdit;
  Num : integer;
begin
  Inherited ;

  For Num := 1 to VH_Paie.PGNbreStatOrg do
  begin
    VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
  end;

  BDelete := TToolBarButton97(GetControl('BDelete'));
  if BDelete <> nil then
    BDelete.OnClick := BDeleteClick;

  MenuVentil := TPopupMenu(GetControl('POPUPVENTIL'));
  if MenuVentil <> nil then
    MenuVentil.Items[0].OnClick := MenuVentilClick;

  Defaut := ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut <> nil then
    Defaut.OnExit := ExitEdit;

  Defaut := ThEdit(getcontrol('PSA_SALARIE_'));
  If Defaut <> nil then
    Defaut.OnExit := ExitEdit;
end ;

procedure TOF_UTOFPGSUPPRANAL.BDeleteClick (Sender: TObject);
var
  F : TFmul;
  i : integer;
  Q : TQuery;
begin
  F := TFmul(Ecran);
  if (F.FListe.nbSelected=0) and (not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné.',F.Caption);
    exit;
  end;

  if GetControlText('RUBRIQUE') = '' then
  begin
    if PGIAsk('Vous n''avez pas saisi de critère Rubrique. #13#10Les ventilations de toutes les rubriques des salariés sélectionnés seront supprimées.#13#10Voulez vous continuer ?', 'Ventilation analytique') <> mrYes then
      exit;
  end
  else
  begin
    Q := OpenSql ('SELECT PRM_RUBRIQUE FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE="'+GetControlText('RUBRIQUE')+'"',TRUE);
    if Q.Eof then
    begin
      if PGIAsk('Vous n''avez pas saisi de critère Rubrique. #13#10Les ventilations de toutes les rubriques des salariés sélectionnés seront supprimées.#13#10Voulez vous continuer ?', 'Ventilation analytique') <> mrYes then
      begin
        Ferme (Q);
        exit;
      end;
    end
    else
    begin
      if (PGIAsk('Voulez vous supprimer les ventilations analytiques salarié par rubrique pour les salariés sélectionnés ?','') <> mrYes) then
        exit;
    end;
    Ferme (Q);
  end;

  if F.FListe.AllSelected then
  begin
  {$IFDEF EAGLCLIENT}
      if not F.FetchLesTous then
        PGIInfo('Impossible de récupérer tous les enregistrements')
      else
  {$ENDIF}
     F.Q.First;
     InitMoveProgressForm(nil,'Suppression des ventilations salarié par rubrique en cours ...', '', (F.Q.recordcount)*2,True,True);
     while not F.Q.Eof do
     begin
        SuppressionVentil;
        MoveCurProgressForm('Suppression');
        F.Q.Next;
     end;
  end
  else
  begin
     InitMoveProgressForm(nil,'Suppression des ventilations salarié par rubrique en cours ...', '', (F.FListe.nbSelected-1)*2,True,True);
     for i := 0 to F.FListe.nbSelected-1 do
     begin
       F.Fliste.GotoLeBookmark(i);
       {$IFDEF EAGLCLIENT}
       F.Q.TQ.Seek(F.FListe.row -1);
       {$ENDIF}
       SuppressionVentil;
       MoveCurProgressForm('Suppression');
     end;
  end;
  FiniMoveProgressForm;

  F.BChercheClick(Ecran);
  F.FListe.ClearSelected;
end;

procedure TOF_UTOFPGSUPPRANAL.SuppressionVentil ();
var
  Rubrique,St : String;
begin
  try
    Rubrique := GetControlText('RUBRIQUE');
    BeginTrans;
    St := ' DELETE FROM VENTIL' +
          ' WHERE V_NATURE LIKE "PG%"' +
          ' AND V_COMPTE LIKE "' + GetField('PSA_SALARIE') + ';' + Rubrique + '%"';
    ExecuteSQL(st);
    CommitTrans;
  except
    Rollback;
    PGIError('Erreur à la suppression des ventilations analytiques du Salarié ' + GetField('PSA_SALARIE') + ' ' + GetField('PSA_LIBELLE') + ' ' + GetField('PSA_PRENOM') + ' - Rubrique ' + Rubrique ,
      'Suppression ventilations');
  end;
end;

procedure TOF_UTOFPGSUPPRANAL.MenuVentilClick(Sender: TObject);
var
  F : TFmul;
  i : integer;
  StSQL : String;
  Q : TQuery;
  Erreur : Boolean;
begin
  // PRINCIPE : on fait la saisie des ventilations sur le premier salarié puis on duplique pour les autres

  // La saisie de la rubrique est obligatoire
  Erreur := False;
  if GetControlText('RUBRIQUE') <> '' then
  begin
    Q := OpenSql ('SELECT PRM_RUBRIQUE FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE="'+GetControlText('RUBRIQUE')+'"',TRUE);
    if Q.Eof then
      Erreur := True;
    Ferme (Q);
  end
  else
    Erreur := True;

  if Erreur then
  begin
    PGIBox('Vous devez sélectionner une rubrique', 'Ventilation analytique');
    exit;
  end;

  F := TFmul(Ecran);
  if (F.FListe.nbSelected=0) and (not F.FListe.AllSelected) then
  begin
    PGIInfo('Aucun élément sélectionné.',F.Caption);
    exit;
  end;

  if (PGIAsk('Voulez vous faire une saisie de masse des ventilations analytiques salarié par rubrique pour les salariés sélectionnés ?','') = mrYes) then
  begin
    if F.FListe.AllSelected then
    begin
    {$IFDEF EAGLCLIENT}
        if not F.FetchLesTous then
          PGIInfo('Impossible de récupérer tous les enregistrements')
        else
    {$ENDIF}
       F.Q.First;
       InitMoveProgressForm(nil,'Saisie des ventilations salarié par rubrique en cours ...', '', (F.Q.recordcount)*2,True,True);
       BeginTrans;
       // Saisie des ventilations pour le premier salarié de la liste
       LeCompte := GetField('PSA_SALARIE') + ';' + GetControlText('RUBRIQUE');
       ParamVentil ('PG',LeCompte , '12345', taModif, False);
       // Récupération des informations saisies pour pouvoir les dupliquer ensuite
       RecupVentilSalarie;
       F.Q.Next;
       while not F.Q.Eof do
       begin
          // Duplication des ventilations pour tous les salariés sélectionnés
          DuplicationVentil;
          MoveCurProgressForm('Saisie');
          F.Q.Next;
       end;
       FreeAndNil(TVentilSal);
       CommitTrans;
    end
    else
    begin
       InitMoveProgressForm(nil,'Saisie des ventilations salarié par rubrique en cours ...', '', (F.FListe.nbSelected-1)*2,True,True);
       BeginTrans;
       F.Fliste.GotoLeBookmark(0);
       // Saisie des ventilations pour le premier salarié de la liste
       LeCompte := GetField('PSA_SALARIE') + ';' + GetControlText('RUBRIQUE');
       ParamVentil ('PG',LeCompte , '12345', taModif, False);
       // Récupération des informations saisies pour pouvoir les dupliquer ensuite
       RecupVentilSalarie;
       for i := 1 to F.FListe.nbSelected-1 do
       begin
         F.Fliste.GotoLeBookmark(i);
         {$IFDEF EAGLCLIENT}
         F.Q.TQ.Seek(F.FListe.row -1);
         {$ENDIF}
         // Duplication des ventilations pour tous les salariés sélectionnés
         DuplicationVentil;
         MoveCurProgressForm('Saisie');
       end;
       FreeAndNil(TVentilSal);
       CommitTrans;
    end;
    FiniMoveProgressForm;
  end;

  F.BChercheClick(Ecran);
  F.FListe.ClearSelected;
end;

procedure TOF_UTOFPGSUPPRANAL.RecupVentilSalarie;
var
  StSQL : String;
  Q : TQuery;
begin
  // Récupérer les valeurs saisies sur le premier salarié
  StSQL := 'SELECT * FROM VENTIL WHERE V_NATURE LIKE "PG%" AND V_COMPTE LIKE "' + LeCompte + '"';
  Q := OpenSql(StSQL,TRUE);
  TVentilSal := TOB.Create('VENTIL', NIL, -1);
  TVentilSal.LoadDetailDB('VENTIL','','',Q,False);
  Ferme(Q);
end;

procedure TOF_UTOFPGSUPPRANAL.DuplicationVentil;
var
  TobVentil,TOBVentilFille : TOB;
  i : integer;
begin
  // Supprimer les éventuelles ventilations salarié/rubrique existantes
  ExecuteSQL('DELETE FROM VENTIL WHERE V_NATURE LIKE "PG%" AND V_COMPTE = "' + GetField('PSA_SALARIE') + ';' + GetControlText('RUBRIQUE') + '"') ;

  // Dupliquer les ventilations
  TOBVentil := TOB.Create ('LESVENTILSDUPLI', nil, -1);
  TOBVentilFille := TOB.Create('VENTIL', TOBVentil, -1);
  for i := 0 to (TVentilSal.Detail.Count -1) do
  begin
    TOBVentilFille := TVentilSal.detail[i];
    TOBVentilFille.PutValue('V_COMPTE', GetField('PSA_SALARIE') + ';' + GetControlText('RUBRIQUE'));
    TOBVentilFille.InsertOrUpdateDB;
  end;
  FreeAndNil(TOBVentil);
end;

procedure TOF_UTOFPGSUPPRANAL.ExitEdit(Sender: TObject);
var
  edit : thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit,10);
end;

Initialization
  registerclasses ( [ TOF_UTOFPGSUPPRANAL ] ) ;
end.
