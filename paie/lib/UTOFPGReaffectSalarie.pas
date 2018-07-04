{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 15/11/2001
Modifié le ... :   /  /
Description .. : Utilitaire de réaffectation du matricule salarié
Mots clefs ... : PAIE;SALARIE
*****************************************************************}
{
PT1    : 02/10/2002 SB V585 FQ n°10249 Mise à jour du paramsoc erronné
PT2    : 17/10/2002 SB V582 Génération d'évènements pour traçage
PT3    : 07/11/2002 SB V585 FQ n° 10317 Rechargement des paramètres sociétés
PT4    : 18/12/2002 SB V591 FQ 10272 Mise à jour des tables VENTIL et VENTANA
PT5    : 01/10/2003 VG V_42 Message "Le controle LIBREAFFECT n'existe pas"
PT6    : 04/06/2004 PH V_50 FQ 10764 controle si numérique et si alpha alors
                            majuscule
PT7    : 06/07/2004 PH V_50 FQ 11405 Prise en compte des profils particuliers du
                            salarié
PT8    : 11/08/2004 VG V_50 Contrôles pour ne pas laisser la possibilité de
                            passer en codage numérique si il existe des
                            matricules alphanumériques - FQ N°11328
PT9    : 27/06/2005 PH V_60 FQ 12109 Requête compatible ORACLE
PT10   : 01/07/2005 SB V_65 FQ 12333 Ajout message d'alerte
PT11   : 10/11/2006 SB V_70 FQ 13383 Nouveau code salarié forcé en majuscule
}
unit UTOFPGReaffectSalarie;

interface
uses StdCtrls, Controls, Classes, Graphics,  sysutils,  ParamSoc,
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HCtrls, HEnt1, HMsgBox, UTOF, UTOB,
  ed_tools ;

type
  TOF_PGREAFFECTSALARIE = class(TOF)
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  private
    Grille: THGrid;
    Tob_Salarie: Tob;
    procedure OnChangeIncrement(Sender: TObject);
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure ClickUtilseInc(Sender: TObject);
  end;

implementation

uses EntPaie, P5Def,PgOutils2;  

{ TOF_PGREAFFECTSALARIE }

procedure TOF_PGREAFFECTSALARIE.ClickUtilseInc(Sender: TObject);
var
  NumEdit: THNumEdit;
begin
  SetControlEnabled('INCREMENT', (GetControlText('UTILISEINC') = 'X'));
  SetControlEnabled('DEBINC', (GetControlText('UTILISEINC') = 'X'));
  NumEdit := THNumEdit(GetControl('INCREMENT'));
  if NumEdit <> nil then
    if GetControlText('UTILISEINC') = 'X' then
      NumEdit.Color := ClWindow
    else
      NumEdit.Color := ClMenu;
  NumEdit := THNumEdit(GetControl('DEBINC'));
  if NumEdit <> nil then
    if GetControlText('UTILISEINC') = 'X' then
      NumEdit.Color := ClWindow
    else
      NumEdit.Color := ClMenu;
end;

procedure TOF_PGREAFFECTSALARIE.GrilleCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var
  T: Tob;
begin
  if Grille.Cells[Acol, Arow] = '' then exit;
  if (ACol = 1) and (VH_Paie.PgTypeNumSal <> 'ALP') then // PT6
  begin
    if IsNumeric(Grille.Cells[Acol, Arow]) then
      Grille.Cells[Acol, Arow] := ColleZeroDevant(StrToInt(Grille.Cells[Acol, Arow]), 10)
    else
    begin
      PgiError('Vous devez saisir une valeur numérique!', Ecran.caption);
      Grille.Cells[Acol, Arow] := Grille.Cells[Acol - 1, Arow];
      exit;
    end;
  end;
  // DEB PT6
  if (VH_Paie.PgTypeNumSal = 'ALP') then Grille.Cells[Acol, Arow] := UPPERCASE(Grille.Cells[Acol, Arow]);
  if (VH_Paie.PgTypeNumSal = 'NUM') and (not IsNumeric(Grille.Cells[Acol, Arow])) then
  begin
    PgiError('Vous devez saisir une valeur numérique!', Ecran.caption);
    Grille.Cells[Acol, Arow] := Grille.Cells[Acol - 1, Arow];
    exit;
  end;
  // FIN P6
  T := Tob_Salarie.FindFirst(['PSA_SALARIE'], ['' + Grille.Cells[Acol, Arow] + ''], False);
  if (T <> nil) and (Grille.Cells[Acol, Arow] <> Grille.Cells[Acol - 1, Arow]) then
  begin
    PgiBox('Vous ne pouvez affecter un matricule existant!', Ecran.caption);
    Grille.Cells[Acol, Arow] := Grille.Cells[Acol - 1, Arow];
  end;
end;

procedure TOF_PGREAFFECTSALARIE.OnArgument(Arguments: string);
var
  i: integer;
  NumEdit: THNumEdit;
  Check: TCheckBox;
begin
  inherited; //ALP : Alphanumérique
  //Num : Numérique
  SetControlText('TYPEMAT', VH_Paie.PgTypeNumSal);
  NumEdit := THNumEdit(GetControl('INCREMENT'));
  if NumEdit <> nil then
  begin
    NumEdit.Masks.PositiveMask := '0';
    NumEdit.OnExit := OnChangeIncrement;
  end;
  NumEdit := THNumEdit(GetControl('DEBINC'));
  if NumEdit <> nil then
  begin
    NumEdit.Masks.PositiveMask := '0';
    NumEdit.OnExit := OnChangeIncrement;
  end;
  Check := TCheckBox(GetControl('UTILISEINC'));
  if Check <> nil then check.OnClick := ClickUtilseInc;
  SetControlChecked('UTILISEINC', False);
  {PT5
  if VH_Paie.PgTypeNumSal='NUM' then
    Begin
    SetControlText('LIBREAFFECT','Réaffectation alphanumérique du matricule');
    End
  Else
    Begin
    SetControlText('LIBREAFFECT','Réaffectation numérique du matricule');
  }
  if VH_Paie.PgTypeNumSal <> 'NUM' then
  begin
    //FIN PT5
    SetControlVisible('LBINC', True);
    SetControlText('INCREMENT', '1');
  end;

  Tob_Salarie := Tob.Create('SALARIES', nil, -1);
  Tob_Salarie.LoadDetailDB('SALARIES', '', '', nil, False);

  Grille := THGrid(GetControl('GDSALARIE'));
  if Grille <> nil then
  begin
    Grille.ColAligns[0] := tacenter;
    Grille.ColAligns[1] := taLeftJustify;
    Grille.ColWidths[0] := 300;
    Grille.ColWidths[1] := 300;
    //Grille.ColFormats[0] := 'CB=PGSALARIE';
    Grille.ColLengths[1] := 10;
    Grille.cells[0, 0] := 'Ancien matricule en ' + LowerCase(RechDom('PGALPHANUM', VH_Paie.PgTypeNumSal, False));
    if (VH_Paie.PgTypeNumSal = 'ALP') then
      Grille.ColFormats[0] := '0000000000';
    for i := 0 to Tob_Salarie.detail.Count - 1 do
    begin
      Grille.cells[0, i + 1] := Tob_Salarie.Detail[i].GetValue('PSA_SALARIE');
      Grille.cells[1, i + 1] := Tob_Salarie.Detail[i].GetValue('PSA_SALARIE');
      if i <> Tob_Salarie.detail.Count then Grille.RowCount := i + 2;
    end;
    Grille.OnCellExit := GrilleCellExit;
  end;

end;

procedure TOF_PGREAFFECTSALARIE.OnChangeIncrement(Sender: TObject);
var
  i, num, Dec, inc: integer;
  St: string;
  T: Tob;
begin
  if GetControlText('INCREMENT') = '0' then
  begin
    for i := 1 to Grille.RowCount do
      Grille.Cells[1, i] := Grille.Cells[0, i];
    Exit;
  end;

  if isnumeric(GetControlText('INCREMENT')) then
    inc := StrtoInt(GetControlText('INCREMENT'))
  else
    inc := 1;
  if isnumeric(GetControlText('DEBINC')) then
    Num := StrtoInt(GetControlText('DEBINC'))
  else
    Num := 0;
  //if VH_Paie.PgTypeNumSal='ALP' then
  for i := 1 to Grille.RowCount do
  begin
    Num := Num + inc;
    St := Collezerodevant(num, 10);
    T := Tob_Salarie.FindFirst(['PSA_SALARIE'], ['' + st + ''], False);
    Dec := num;
    while T <> nil do
    begin
      Dec := Dec + 1;
      if Dec = Num + inc then Num := Dec;
      St := Collezerodevant(Dec, 10);
      T := Tob_Salarie.FindNext(['PSA_SALARIE'], ['' + st + ''], False);
    end;
    Grille.Cells[1, i] := st;
  end;
end;

procedure TOF_PGREAFFECTSALARIE.OnClose;
begin
  inherited;
  if Tob_Salarie <> nil then
  begin
    Tob_Salarie.free;
    Tob_Salarie := nil
  end;
end;

procedure TOF_PGREAFFECTSALARIE.OnUpdate;
var
q: TQuery;
Tob_Table, T, Tob_Ventil: Tob;
AncSal, NewSal, Anc_VCompte, New_VCompte, Nomchamp, NomTable, St: string;
i, j: integer;
TEvent: TStringList; //PT2
NumPossible, Temp: Boolean;
begin
inherited;
{ DEB PT11 }
i := Grille.col;
j := Grille.Row;
GrilleCellExit(Grille,i,j,Temp);
{ FIN PT11 }

NumPossible:= True;  //PT8
Tob_Ventil := nil;
if PgiAsk('Etes-vous sûr de vouloir lancer la mise à jour ?', Ecran.Caption) = mrNo then
   exit;
InitMoveProgressForm(nil, 'Chargement des données',
                     'Veuillez patienter SVP ...', Grille.RowCount, FALSE, TRUE);

//InitMove(Grille.RowCount,'');
TEvent := TStringList.create; //PT2
q:= OpenSql ('SELECT DT_NOMTABLE, DT_LIBELLE, DH_NOMCHAMP, DT_DOMAINE'+
             ' FROM DECHAMPS'+
             ' LEFT JOIN DETABLES ON'+
             ' DH_PREFIXE=DT_PREFIXE WHERE'+
             ' DH_NOMCHAMP LIKE "%_SALARIE" AND'+
             ' DH_TYPECHAMP LIKE "VARCHAR%"'+ //PT1 Ajout clause
             ' ORDER BY DT_DOMAINE, DT_NOMTABLE', True);
Tob_Table:= Tob.create('Liste des tables', nil, -1);
Tob_Table.LoadDetailDB('DECHAMPS', '', '', Q, True);
Ferme(Q);

try
   BeginTrans;
   for i := 1 to Grille.RowCount do
       begin
       MoveCurProgressForm('Mise à jour salarié : ' + Grille.cells[0, i]);
{PT8
       if (Grille.cells[0, i]<>Grille.cells[1, i]) and
          (Grille.cells[1, i]<>'') then
          begin
          AncSal:= Grille.cells[0, i];
          NewSal:= Grille.cells[1, i];
}
       AncSal:= Grille.cells[0, i];
       NewSal:= Grille.cells[1, i];

       if (not (IsNumeric(NewSal))) then
          NumPossible:= False;

       if (AncSal<>NewSal) and (NewSal<>'') then
          begin
//FIN PT8
          TEvent.Add ('Réaffectation matricule salarié '+AncSal+
                      ' nouveau matricule '+NewSal); //PT2
          if VH_PAie.PGCodeInterim then
             if ExisteSql ('SELECT PSI_INTERIMAIRE'+
                           ' FROM INTERIMAIRES WHERE'+
                           ' PSI_INTERIMAIRE="'+NewSal+'" ') then
                begin
                PGIBox ('Réaffectation impossible pour ce salarié.#13#10'+
                        'Un intérimaire existe avec le même nouveau matricule.',
                        'Matricule existant : '+NewSal);
                Continue;
                end;
//DEB PT4 Chargement des enregistrements de la table VENTIL avant réaffectation
//table SALARIES
// DEB PT9  Requête revue pour la rendre compatible ORACLE.
          Q:= OpenSql ('SELECT V_NATURE,V_COMPTE,V_NUMEROVENTIL'+
                       ' FROM VENTIL,SALARIES'+
                       ' WHERE (V_COMPTE LIKE PSA_SALARIE||"%") AND'+
                       ' V_NATURE LIKE "PG%" AND'+
                       ' PSA_SALARIE="'+AncSal+'"'+
                       ' ORDER BY V_COMPTE,V_NATURE ', True);
// FIN PT9
          if not Q.Eof then
             begin
             Tob_Ventil:= TOB.Create('Les ventilations', nil, -1);
             Tob_Ventil.LoadDetailDB('Les ventilations', '', '', Q, False);
             end;
          Ferme(Q);
//FIN PT4
          T:= Tob_Table.FindFirst([''], [''], False);
          while T <> nil do
                begin
                NomChamp:= T.GetValue('DH_NOMCHAMP');
                if Trim(Copy(NomChamp, 5, Length(NomChamp))) = 'SALARIE' then
                   begin
                   NomTable:= T.GetValue('DT_NOMTABLE');
                   MoveCurProgressForm( 'Salarié : '+AncSal+' mise à jour '+
                                        RechDom('TTTABLES', NomTable, False));
                   TEvent.Add ('Traitement de la table '+
                               Trim(RechDom('TTTABLES', NomTable, False))+
                               ' pour le salarié '+NewSal); //PT2
                   ExecuteSql ('UPDATE '+NomTable+' SET '+
                               NomChamp+'="'+NewSal+'" WHERE '+
                               NomChamp+'="'+AncSal+'"');
                   end;
                T := Tob_Table.FindNext([''], [''], False);
                end;
{DEB PT4 MAJ DES TABLES INTERIMAIRES, EMPLOIINTERIM,VENTIL ET VENTANA}
          if VH_PAie.PGCodeInterim then
             begin
             MoveCurProgressForm ('Salarié : '+AncSal+' mise à jour '+
                                  RechDom('TTTABLES', 'INTERIMAIRES', False));
             if not ExisteSql ('SELECT PSI_INTERIMAIRE'+
                               ' FROM INTERIMAIRES WHERE'+
                               ' PSI_INTERIMAIRE="'+NewSal+'"') then
                begin
                ExecuteSql ('UPDATE INTERIMAIRES SET'+
                            ' PSI_INTERIMAIRE="'+NewSal+'" WHERE'+
                            ' PSI_INTERIMAIRE="'+AncSal+'"');
                MoveCurProgressForm ('Salarié : '+AncSal+' mise à jour '+
                                     RechDom('TTTABLES', 'EMPLOIINTERIM', False));
                ExecuteSql ('UPDATE EMPLOIINTERIM SET'+
                            ' PEI_INTERIMAIRE="'+NewSal+'" WHERE'+
                            ' PEI_INTERIMAIRE="'+AncSal+'"');
                end
             else
                PGIInfo('Un intérimaire existe avec le même nouveau code.');
             end;
          MoveCurProgressForm ('Salarié : '+AncSal+' mise à jour '+
                               RechDom('TTTABLES', 'VENTIL', False));
// DEB PT7 Mise à jour des champs de la table PROFILSPECIAUX
          ExecuteSql ('UPDATE PROFILSPECIAUX SET'+
                      ' PPS_CODE="'+NewSal+'" WHERE'+
                      ' PPS_CODE="'+AncSal+'" AND'+
                      ' PPS_ETABSALARIE="-" ');
          MoveCurProgressForm ('Profils particuliers du salarié : '+AncSal+
                               ' traités');
// FIN PT7
//Mise à jour des champs V_COMPTE=PSA_SALARIE nature SA
          ExecuteSql ('UPDATE VENTIL SET'+
                      ' V_COMPTE="'+NewSal+'" WHERE'+
                      ' V_COMPTE="'+AncSal+'" AND'+
                      ' V_NATURE LIKE "SA%"');

//Mise à jour des champs V_COMPTE like PSA_SALARIE nature PG
          if Tob_Ventil <> nil then
             begin
             TEvent.Add ('Traitement de la table '+
                         Trim(RechDom('TTTABLES', 'VENTIL', False))+
                         ' pour le salarié '+NewSal);
             for j := 0 to Tob_Ventil.detail.count - 1 do
                 begin
                 MoveCurProgressForm ('Salarié : '+AncSal+' mise à jour '+
                                      RechDom('TTTABLES', 'VENTANA', False));
                 T:= Tob_Ventil.detail[j];
                 Anc_VCompte:= T.GetValue('V_COMPTE');
                 ReadTokenSt(Anc_VCompte);
                 New_VCompte:= NewSal+';'+Anc_VCompte;
                 St:= 'UPDATE VENTIL SET'+
                      ' V_COMPTE="'+New_VCompte+'" WHERE'+
                      ' V_COMPTE="'+T.GetValue('V_COMPTE')+'" AND'+
                      ' V_NATURE="'+T.GetValue('V_NATURE')+'" AND'+
                      ' V_NUMEROVENTIL='+IntToStr(T.GetValue('V_NUMEROVENTIL'));
                 ExecuteSql(st);
                 end;
             end;

          if Tob_Ventil <> nil then
             FreeAndNil(Tob_Ventil);

//Mise à jour de la table Ventana
          Q:= OpenSql ('SELECT YVA_TABLEANA,YVA_IDENTIFIANT,YVA_NATUREID,'+
                       ' YVA_IDENTLIGNE,YVA_AXE,YVA_NUMVENTIL'+
                       ' FROM VENTANA WHERE'+
                       ' YVA_NATUREID="PG" AND'+
                       ' YVA_IDENTIFIANT LIKE "'+AncSal+';%"', True);
          if not Q.Eof then
             begin
             Tob_Ventil:= TOB.Create('Les ventilations', nil, -1);
             Tob_Ventil.LoadDetailDB('Les ventilations', '', '', Q, False);
             Ferme(Q);
             TEvent.Add ('Traitement de la table '+
                         Trim(RechDom('TTTABLES', 'VENTANA', False))+
                         ' pour le salarié '+NewSal);
             for j := 0 to Tob_Ventil.detail.count - 1 do
                 begin
                 MoveCurProgressForm('Mise à jour des pré-ventilations analytiques..');
                 T:= Tob_Ventil.detail[j];
                 Anc_VCompte:= T.GetValue('YVA_IDENTIFIANT');
                 ReadTokenSt(Anc_VCompte);
                 New_VCompte:= NewSal+';'+Anc_VCompte;
                 St:= 'UPDATE VENTANA SET'+
                      ' YVA_IDENTIFIANT="'+New_VCompte+'" WHERE'+
                      ' YVA_IDENTIFIANT="'+T.GetValue('YVA_IDENTIFIANT')+'" AND'+
                      ' YVA_TABLEANA="'+T.GetValue('YVA_TABLEANA')+'" AND'+
                      ' YVA_NATUREID="'+T.GetValue('YVA_NATUREID')+'" AND'+
                      ' YVA_IDENTLIGNE="'+T.GetValue('YVA_IDENTLIGNE')+'" AND'+
                      ' YVA_AXE="'+T.GetValue('YVA_AXE')+'" AND'+
                      ' YVA_NUMVENTIL='+IntToStr(T.GetValue('YVA_NUMVENTIL'));
                 ExecuteSql(st);
                 end;
             end
          else
             Ferme(Q);
          if Tob_Ventil <> nil then
             FreeAndNil(Tob_Ventil);
{FIN PT4}
          end;
      end;

//mise à jour du type matricule ds PARAMSOC
   if VH_Paie.PgTypeNumSal = 'ALP' then
      begin
//PT8
      if (NumPossible=True) then
         begin
         if PgiAsk ('Voulez-vous modifier le paramètrage société donc gérer le'+
                    ' matricule en numérique ?', Ecran.Caption) = mrYes then
            begin //PT1 mise en commentaire
//  ExecuteSql('UPDATE PARAMSOC SET SOC_DATA="NUM" WHERE SOC_NOM="SO_PGTYPENUMSAL"');
            MoveCurProgressForm('Mise des paramètres sociétés..');
            SetParamSoc('SO_PGTYPENUMSAL', 'NUM');
            ChargeParamsPaie;
            ; //PT3
            end;
         end;
//FIN PT8         
      end
   else
      begin
      if PgiAsk ('Voulez-vous modifier le paramètrage société donc gérer le'+
                 ' matricule en alphanumérique ?', Ecran.Caption) = mrYes then
      { DEB PT10 }
      if PgiAsk ('Attention! Si vous confirmez le passage des matricules salariés en gestion alphanumérique,#13#10'+
      'vous ne pourrez plus revenir en gestion numérique.#13#10'+
      'L''incrémentation automatique des matricules ne sera plus possible.#13#10'+
      'Les auxiliaires salariés ne pourront plus être créés automatiquement. Confirmez-vous ?"') = mrYes then
      { FIN PT10 }
         begin //PT1 Mise en commentaire et modif setparamsoc
//    ExecuteSql('UPDATE PARAMSOC SET SOC_DATA="ALP" WHERE SOC_NOM="SO_PGTYPENUMSAL"');
         MoveCurProgressForm('Mise des paramètres sociétés..');
         SetParamSoc('SO_PGTYPENUMSAL', 'ALP');
         ChargeParamsPaie;
         ; //PT3
         end;
      end;
      MoveCurProgressForm('Enregistrement des données dans la base..patientez..');
   CommitTrans;
   TEvent.Add('La mise à jour s''est correctement terminée.'); //PT2
   CreeJnalEvt('004', '124', 'OK', nil, nil, TEvent); //PT2
   PGIBox ('La mise à jour s''est correctement terminée. Fin du traitement',
           Ecran.caption);
except
   Rollback; //PT2
   TEvent.Add ('Une erreur est survenue lors de la mise à jour du matricule'+
               ' salarié : '+AncSal);
   CreeJnalEvt('004', '124', 'ERR', nil, nil, TEvent); //PT2
   PGIBox ('Une erreur est survenue lors de la mise à jour du matricule'+
           ' salarié : '+AncSal, Ecran.caption);
   end;

MoveCurProgressForm('Chargement des nouveaux matricules..');

//Button8Click(nil);
//FiniMove;
FiniMoveProgressForm;
FreeAndNil(Tob_Table);
FreeAndNil(Tevent); //PT2
end;

initialization
  registerclasses([TOF_PGREAFFECTSALARIE]);
end.

