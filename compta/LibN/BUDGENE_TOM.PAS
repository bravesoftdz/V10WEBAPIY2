{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 10/01/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BUDGENE (BUDGENE)
Suite ........ : 
Suite ........ : JP 11/08/05 : FQ 16044 : On renseigne le code rubrique 
Suite ........ : par défaut avec le code section
Mots clefs ... : TOM;BUDGENE
*****************************************************************}
Unit BUDGENE_TOM;

Interface

Uses Windows, StdCtrls, Controls, Classes, Graphics, Buttons,
{$IFDEF EAGLCLIENT}
     eFiche, eFichList, Maineagl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables {$ELSE} uDbxDataSet {$ENDIF},
     HREGCPTE, HDB, db, Fiche, FichList, FE_Main,
{$ENDIF}
     PLANBUDGET_TOF, // PlanBudget
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, HTB97, UTOM, utilPGI, AglInit, Grids,
     UTob;

procedure FicheBudgene(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
procedure FicheBudgeneMZS(Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);

Type
  TOM_BUDGENE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    Fiche : TFFiche;
    FListe : THGrid;
    MsgBox : THMsgBox;
    Mode : TActionFiche;
    LaPage : Integer;
    LesModif, Lequel : String;
    SensDefaut,SigneDefaut : String;
    LgMin,LgMax : Integer ;
    procedure BImprimerClick(Sender: TObject);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    {JP 11/08/05 : FQ 16044 : Pour renseigner le code rubrique par défaut}
    procedure CodeOnExit(Sender : TObject);

    function Grid2St ( ACol : Integer ) : String;
    procedure LitCompteExclu(St : String ; ACol : Integer);
    procedure EcritCompteExclu;
    procedure RempliGrid;
    procedure AffecteLe(Champ, Valeur : string);
    procedure InitModifEnSerie(StModif : string);
    function CodeValide : Boolean;
    procedure ChargeLgMinMax;
    procedure CurseurAuCode;
    procedure CurseurAuCodeRub;
    function PresenceRub : Boolean;
    function OkConfidentiel : Boolean;
    procedure UpdateDate;
  end;

function EstMouvementeBudgen(const LeCpte : String) : Boolean;
function ExisteBudGen(const LeCpte : String) : Boolean;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  {$ENDIF MODENT1}
  Paramsoc, Ent1;


procedure FicheBudgene(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer) ;
begin
  if ((Comment in [taCreat..taCreatOne])=False) then if not ExisteBudGen(Compte) then exit;
  if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit;
  AGLLanceFiche('CP', 'CPBUDGENE', '', Compte, ActionToString(Comment)+';'+Axe+';'+Compte+';'+IntToStr(QuellePage));
end;

procedure FicheBudgeneMZS(Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
begin
  if ((Comment in [taCreat..taCreatOne])=False) then if not ExisteBudGen(Lequel) then exit;
  if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit;
  AGLLanceFiche('CP', 'CPBUDGENE', '', Lequel, TAToStr(Comment)+';;'+Lequel+';'+IntToStr(QuellePage)+';'+LesModif);
end;

function EstMouvementeBudgen(const LeCpte : String) : Boolean;
begin
  Result := ExisteSQL('SELECT BG_BUDGENE FROM BUDGENE WHERE BG_BUDGENE="'+LeCpte+'" '+
              'AND (EXISTS(SELECT BE_BUDGENE FROM BUDECR WHERE BE_BUDGENE="'+LeCpte+'"))');
end;

function ExisteBudGen(const LeCpte : String) : Boolean;
begin
  Result := True;
  if LeCpte='' then exit;
  if not ExisteSQL('SELECT BG_BUDGENE FROM BUDGENE WHERE BG_BUDGENE="'+LeCpte+'"') then begin
    MessageAlerte('L''enregistrement est inaccessible.');
    Result := False;
  end;
end;

procedure TOM_BUDGENE.OnNewRecord ;
begin
  Inherited ;

  UpdateDate;

  SetField('BG_SENS', 'M');
  SetControlChecked('BG_HT.Checked', False);
  SetControlEnabled('BG_BUDGENE', True);

  SetFocusControl('BG_BUDGENE');
  if Mode=taCreatEnSerie Then begin
    if SensDefaut<>'' then SetField('BG_SENS', SensDefaut);
    if SigneDefaut<>'' then SetField('BG_SIGNE', SigneDefaut);
  end;
end ;

procedure TOM_BUDGENE.OnDeleteRecord ;
begin
  Inherited;

  if EstMouvementeBudgen(GetField('BG_BUDGENE')) then begin
    LastError := 1;
    LastErrorMsg := MsgBox.Mess.Strings[1]; // Ce compte budgétaire comporte des écritures comptables, vous ne pouvez pas le supprimer.
    Exit;
  end;
end ;

procedure TOM_BUDGENE.OnUpdateRecord ;
begin
  Inherited ;
  if Not CodeValide then begin
    LastError :=1;
    exit;
  end;

  if DS.State in [dsEdit,dsInsert] then begin
   if GetField('BG_RUB') = '' then begin
     Msgbox.Execute(5,'',''); // Vous devez renseigner le code de rubrique.
     CurseurAuCodeRub;
     LastError :=1;
     Exit;
   end;
   if FListe.Cells[0,1]='' then begin
     if MsgBox.Execute(6,'','')<>mrYes then begin // Vous n''avez pas associé de comptes généraux à votre compte budgétaire. Désirez-vous continuer ?
       Fiche.Pages.ActivePage := Fiche.Pages.Pages[1];
       FListe.SetFocus;
       LastError :=1;
       Exit;
     end;
   end;
   if PresenceRub then begin
     if MsgBox.Execute(7,'','')<>mrYes then begin // Le code rubrique que vous avez renseigné existe déjà. La génération automatique des rubriques ne sera pas complète. Désirez-vous continuer ?
       CurseurAuCodeRub;
       LastError :=1;
       Exit;
     end;
   end;
  end;
  EcritCompteExclu;
end ;

procedure TOM_BUDGENE.OnAfterUpdateRecord ;
begin
  Inherited ;

  if mode = taCreatEnSerie then begin
    SensDefaut := GetField('BG_SENS');
    SigneDefaut := GetField('BG_SIGNE');
  end;
end ;

procedure TOM_BUDGENE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BUDGENE.OnLoadRecord ;
begin
  Inherited ;

  UpdateDate;

  case Mode Of
    taConsult : begin
      FListe.Enabled := True;
      FListe.Options := FListe.Options-[goEditing]+[goRowSelect];
    end;
    taModifEnSerie : InitModifEnSerie(LesModif);
  end;
  RempliGrid;
end ;

procedure TOM_BUDGENE.OnChangeField ( F: TField ) ;
begin
  Inherited ;

 {JP 11/08/05 : FQ 16044 : Déplacer dans CodeOnExit, car ici, c'est trop tôt
  if ((F.FieldName='BG_BUDGENE') and (GetField('BG_RUB')='')) then
    SetField('BG_RUB', Copy(F.AsString,1,5));}
end ;

// 1° : TActionFiche
// 2° : Axe (non utilisé)
// 3° : Lequel ou Compte
// 4° : QuellePage
// 5° : LesModif
procedure TOM_BUDGENE.OnArgument ( S: String ) ;
var
  shtLibre : TTabSheet;
begin
  Inherited ;

  // Paramètres
  Mode := StrToTA(ReadTokenST(S));
  ReadTokenST(S); // Axe
  Lequel := ReadTokenST(S);
  LaPage := READTOKENI(S);
  LesModif := ReadTokenST(S);

  MsgBox := THMsgBox.Create(Ecran);
  MsgBox.Mess.Add('0;Comptes budgétaires;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
  MsgBox.Mess.Add('Ce compte budgétaire comporte des écritures comptables, vous ne pouvez pas le supprimer.');
  MsgBox.Mess.Add('2;Comptes budgétaires;Votre saisie des comptes est trop longue, elle sera tronquée.;E;O;O;O;');
  MsgBox.Mess.Add('3;Comptes budgétaires;La longueur du code que vous avez saisi est inférieure à la longueur définie pour cette société.;W;O;O;O;');
  MsgBox.Mess.Add('4;Comptes budgétaires;La longueur du code que vous avez saisi est supérieure à la longueur définie pour cette société.;W;O;O;O;');
  MsgBox.Mess.Add('5;Comptes budgétaires;Vous devez renseigner le code de rubrique.;W;O;O;O;');
  MsgBox.Mess.Add('6;Comptes budgétaires;Vous n''avez pas associé de comptes généraux à votre compte budgétaire. Désirez-vous continuer ?;Q;YNC;N;N;');
  MsgBox.Mess.Add('7;Comptes budgétaires;Le code rubrique que vous avez renseigné existe déjà. La génération automatique des rubriques ne sera pas complète. Désirez-vous continuer ?;Q;YNC;N;N;');
  MsgBox.Mess.Add('L''accès à ce compte est interdit.');

  shtLibre := TTabSheet(GetControl('ZL', True));
  if (shtLibre <> nil) then LibellesTableLibre(shtLibre, 'TB_TABLE', 'BG_TABLE', 'B');

  Fiche := TFFiche(Ecran);
  FListe := THGrid(GetControl('FLISTE', True));
  FListe.OnKeyPress := FListeKeyPress;
  FListe.OnKeyDown := FListeKeyDown;
  FListe.OnCellExit := FListeCellExit;
  Fiche.BImprimer.Visible := True;
  Fiche.BImprimer.OnClick := BImprimerClick;

  {JP 11/08/05 : FQ 16044 : Pour renseigner le code rubrique par défaut}
  THEdit(GetControl('BG_BUDGENE', True)).OnExit := CodeOnExit;

  ChargeLgMinMax;
  if Not OkConfidentiel then Exit;
  Fiche.Pages.ActivePage := Fiche.Pages.Pages[LaPage];
  case Mode Of
    taConsult : begin
      FListe.Enabled := True;
      FListe.Options := FListe.Options-[goEditing]+[goRowSelect];
    end;
    taCreat..taCreatOne : begin
      SetField('BG_BUDGENE', Lequel);
      Fiche.bDefaire.Enabled:=False;
    end;
  end;
  SensDefaut := '';
  SigneDefaut := '';
  Ecran.HelpContext := 15113100;
end ;

procedure TOM_BUDGENE.OnClose ;
begin
  Inherited ;

  MsgBox.Free;
end ;

procedure TOM_BUDGENE.OnCancelRecord ;
begin
  Inherited ;
end ;

function TOM_BUDGENE.Grid2St(ACol: Integer): String;
var
  i : Integer;
  St : String;
begin
  St:='';
  for i:=1 to FListe.RowCount-1 do
    St := St + FListe.Cells[ACol,i]+';';
  Result := St;
  if Length(Result)>250 then MsgBox.Execute(2, '', ''); // Votre saisie des comptes est trop longue, elle sera tronquée.
end;

procedure TOM_BUDGENE.LitCompteExclu(St: String; ACol: Integer);
var
  i : Integer;
  St1 : String;
begin
  i:=1;
  while St <> '' do begin
    St1:=ReadTokenSt(St);
    if St1 <> '' then begin
      FListe.Cells[ACol,i]:=St1;
      Inc(i);
      end
    else begin
      if ACol=1 then Inc(i);
    end;
    if i > Fliste.RowCount-1 then FListe.RowCount:=FListe.RowCount+1;
  end;

  if FListe.Cells[0,FListe.RowCount-2]='' then FListe.RowCount:=FListe.RowCount-1;
end;

procedure TOM_BUDGENE.EcritCompteExclu;
begin
  SetField('BG_COMPTERUB', Grid2St(0));
  SetField('BG_EXCLURUB', Grid2St(1));
end;

procedure TOM_BUDGENE.RempliGrid;
begin
  FListe.VidePile(False);
  LitCompteExclu(GetField('BG_COMPTERUB'),0);
  LitCompteExclu(GetField('BG_EXCLURUB'),1);
end;

procedure TOM_BUDGENE.AffecteLe(Champ, Valeur: string);
var
  C : TControl;
begin
  C := GetControl(Champ);

  {$IFDEF EAGLCLIENT}
  if (C is TCheckBox) or (C is THValComboBox) or (C is TEdit) or (C is THEdit) then
  {$ELSE}
  if (C is THDBCheckBox) or (C is THDBValComboBox) or (C is THDBEdit) or (C is THDBCpteEdit) then
  {$ENDIF}
    begin
    SetField(Champ, Valeur);
    TEdit(C).Font.Color := clRed;
    end
  {$IFDEF EAGLCLIENT}
  else if (C is THSpinEdit) then
  {$ELSE}
  else if (C is THDBSpinEdit) then
  {$ENDIF}
    begin
    SetField(Champ,Valeur);
    THSpinEdit(C).Font.Color := clRed;
  end;
end;

procedure TOM_BUDGENE.InitModifEnSerie(StModif: string);
var
  St,Champ, Valeur : string;
  i : integer;
  B : TBitBtn;
  PBouton : TToolWindow97;
begin
  if DS.State = dsBrowse then DS.Edit;

  While StModif<>'' do begin
    St:=ReadTokenSt(StModif);
    i:=Pos('=',St); if i>0 then Champ := Trim(Copy(St,1,i-1));
    i:=Pos('"',St); if i>0 then St := Trim(Copy(St,i+1,Length(St)));
    i:=Pos('"',St); if i>0 then Valeur := Trim(Copy(St,1,i-1));
    AffecteLe(Champ,Valeur);
  end;

  PBouton := TToolWindow97(GetControl('PBOUTON', True));
  for i:=0 to PBouton.ControlCount-1 do begin
    if PBouton.Controls[i] is TBitBtn then begin
      B := TBitBtn(PBouton.Controls[i]);
      if ((UpperCase(B.Name)<>'BVALIDER') and
          (UpperCase(B.Name)<>'BFERME') and
          (UpperCase(B.Name)<>'BHELPBTN')) then B.Enabled := False;
    end;
  end;
end;

function TOM_BUDGENE.CodeValide: Boolean;
var
  L : Integer;
begin
  Result := False;
  L := Length(GetField('BG_BUDGENE'));
  if DS.state in [dsInsert] then begin
    if L<LgMin then begin
      CurseurAuCode;
      Msgbox.Execute(3, '', ''); // La longueur du code que vous avez saisi est inférieure à la longueur définie pour cette société.
      Exit;
      end else
    if L>LgMax then begin
      CurseurAuCode;
      Msgbox.Execute(4, '', ''); // La longueur du code que vous avez saisi est supérieure à la longueur définie pour cette société.
      Exit;
      end else
    if Presence('BUDGENE','BG_BUDGENE',GetField('BG_BUDGENE'))  then begin
      CurseurAuCode;
      MsgBox.Execute(0, '',''); // Le code que vous avez saisi existe déjà. Vous devez le modifier.
      Exit;
    end;
  end;
  Result := True;
end;

procedure TOM_BUDGENE.ChargeLgMinMax;
begin
  LgMax := GetParamsocSecur('SO_LGMAXBUDGET',0);
  LgMin := GetParamsocSecur('SO_LGMINBUDGET',0);
  if LgMax=0 then LgMax:=17;
end;

procedure TOM_BUDGENE.CurseurAuCode;
begin
  Fiche.Pages.ActivePage := Fiche.Pages.Pages[0];
  SetFocusControl('BG_BUDGENE');
end;

procedure TOM_BUDGENE.CurseurAuCodeRub;
begin
  Fiche.Pages.ActivePage := Fiche.Pages.Pages[0];
  SetFocusControl('BG_RUB');
end;

function TOM_BUDGENE.PresenceRub: Boolean;
begin
  if Length(GetField('BG_RUB'))>5 then SetField('BG_RUB', Copy(GetField('BG_RUB'),1,5));
  Result := PresenceComplexe('BUDGENE',['BG_BUDGENE','BG_RUB'],['<>','='],[GetField('BG_BUDGENE'),GetField('BG_RUB')],['S','S']);
end;

procedure TOM_BUDGENE.BImprimerClick(Sender: TObject);
begin
{$IFNDEF IMP}
PlanBudget(GetField('BG_BUDGENE'), True);
{$ENDIF}
end;

procedure TOM_BUDGENE.FListeKeyPress(Sender: TObject; var Key: Char);
begin
  if Mode=taConsult then Exit;
  if Not (DS.State in [dsEdit,dsInsert]) then begin
    DS.Edit;
//    QBudgeneBG_LIBELLE.AsString:=QBudgeneBG_LIBELLE.AsString ;
  end;
  Key := UpCase(Key);
end;

procedure TOM_BUDGENE.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Mode = taConsult then Exit;
  if Not(ssCtrl in Shift) then Exit;
  if Key<>VK_DELETE then Exit;
  if FListe.RowCount<=2 then Exit;

  FListe.DeleteRow(FListe.Row);
  if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount:=FListe.RowCount+1;
  if not(DS.Modified) then begin
    if not(DS.State in [dsEdit,dsInsert]) then DS.Edit;
//    QBudgeneBG_LIBELLE.AsString:=QBudgeneBG_LIBELLE.AsString ;
  end;
end;

procedure TOM_BUDGENE.FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Mode = taConsult then Exit;
  if DS.Modified then
    if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount := FListe.RowCount+1;
end;

function TOM_BUDGENE.OkConfidentiel: Boolean;
begin
  Result:=False ;
  if V_PGI.Confidentiel='0' then begin
    if (GetCheckBoxState('BG_CONFIDENTIEL') = cbChecked) then begin
      MessageAlerte(MsgBox.Mess[8]); // L''accès à ce compte  est interdit.
      Close;
      Exit;
    end;
  end;
  SetControlVisible('BG_CONFIDENTIEL', (V_PGI.Confidentiel = '1'));
  SetControlEnabled('BG_CONFIDENTIEL', V_PGI.Superviseur);
  Result := True;
end;

procedure TOM_BUDGENE.UpdateDate;
var
  szDate : String;
begin
  szDate := FormatDateTime('dd mmm yyyy', GetField('BG_DATECREATION'));
  SetControlText('DATECREATION', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BG_DATEMODIF'));
  SetControlText('DATEMODIF', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BG_DATEOUVERTURE'));
  SetControlText('DATEOUVERTURE', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BG_DATEFERMETURE'));
  SetControlText('DATEFERMETURE', szDate);
end;

{JP 11/08/05 : FQ 16044 : Pour renseigner le code rubrique par défaut
{---------------------------------------------------------------------------------------}
procedure TOM_BUDGENE.CodeOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if (GetField('BG_RUB') = '') and (DS.State in [dsInsert, dsEdit]) then
    SetField('BG_RUB', Copy(GetField('BG_BUDGENE'), 1, 5));
end;

Initialization
  registerclasses ( [ TOM_BUDGENE ] ) ;
end.
