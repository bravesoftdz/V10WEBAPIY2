{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 11/01/2005
Modifié le ... : 25/01/2006
Description .. : Source TOM de la TABLE : BUDSECT (BUDSECT)
Suite ........ : 
Suite ........ : JP 11/08/05 : FQ 16045 : On renseigne le code rubrique
Suite ........ : par défaut avec le code section
Suite ........ : 
Suite ........ : Lek, je compile en mode 2/3:
Suite ........ : pour être compilé j'ajoute
Suite ........ : dbCtrls, hRegCpte, hDb dans le mode 2/3
Mots clefs ... : TOM;BUDSECT
*****************************************************************}
Unit BUDSECT_TOM ;

Interface

Uses Windows, StdCtrls, Controls, Classes, Graphics, Buttons,
     PLANBUDSEC_TOF, // PlanBudSec
{$IFDEF EAGLCLIENT}
     eFiche, eFichList, Maineagl,
{$ELSE}
     dbCtrls, hRegCpte,hDb, {Lek 250106 pour compiler en 2/3}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche, FichList, FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, HTB97, UTOM, utilPGI, AglInit, Grids,
     Hcompte,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     Ent1,
     UTob;

procedure FicheBudsect(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
procedure FicheBudsectMZS(Axe,Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
function SectionRetrieBud(CodRupt : String ; Axe : String ; ListeCodes : TStringList) : TSectRetri;

Type
  TOM_BUDSECT = Class (TOM)
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
    AxeDefaut, QuelAxe, AxeSerie, MemoAxe, SensDefaut, SigneDefaut : String;
    procedure BImprimerClick(Sender: TObject);
    procedure FListeKeyPress(Sender: TObject; var Key: Char);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeCellExit(Sender: TObject; var ACol, ARow: Longint; var Cancel: Boolean);
    procedure BS_AXEChange(Sender: TObject);
    {JP 11/08/05 : FQ 16045 : Pour renseigner le code rubrique par défaut}
    procedure CodeOnExit(Sender : TObject);

    function Grid2St ( ACol : Integer ) : String;
    procedure LitCompteExclu(St : String ; ACol : Integer);
    procedure EcritCompteExclu;
    procedure RempliGrid;
    procedure AffecteLe(Champ, Valeur : string);
    procedure InitModifEnSerie(StModif : string);
    function CodeValide : Boolean;
    procedure CurseurAuCode;
    procedure CurseurAuCodeRub;
    function PresenceRub : Boolean;
    function OkConfidentiel : Boolean;
    procedure UpdateDate;
  end;

function EstMouvementeBudSect(const LeCpte,Laxe : String) : Boolean;
function ExisteBudSect(const LeCpte, Laxe : String) : Boolean;


Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  {$ENDIF MODENT1}
  Paramsoc;

procedure FicheBudsect(Q : TQuery ; Axe,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
var
  ch : string; {JP 03/08/05 : FQ 16058}
begin
  if ((Comment in [taCreat..taCreatOne])=False) then if not ExisteBudSect(Compte, Axe) then exit;
  if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit;

  {JP 03/08/05 : FQ 16058 : pour éviter si Axe et Compte sont vides en création que le code soit initialisé avec ';'}
  if Axe <> '' then begin
    ch := Axe + ';';
    if Compte <> '' then
      ch := ch + Compte + ';';
  end;

  AGLLanceFiche('CP', 'CPBUDSECT', '', ch, ActionToString(Comment)+';'+Axe+';'+Compte+';'+IntToStr(QuellePage));
end;

procedure FicheBudsectMZS(Axe,Lequel : String ; Comment : TActionFiche ; QuellePage : Integer; LesModif : string);
var
  ch : string; {JP 03/08/05 : FQ 16058}
begin
  if ((Comment in [taCreat..taCreatOne])=False) then if not ExisteBudSect(Lequel, Axe) then exit;
  if _Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit;

  {JP 03/08/05 : FQ 16058 : pour éviter si Axe et Compte sont vides en création que le code soit initialisé avec ';'}
  if Axe <> '' then begin
    ch := Axe + ';';
    if Lequel <> '' then
      ch := ch + Lequel + ';';
  end;

  AGLLanceFiche('CP', 'CPBUDSECT', '', Axe+';'+Lequel, TAToStr(Comment)+';'+Axe+';'+Lequel+';'+IntToStr(QuellePage)+';'+LesModif);
end;

function SectionRetrieBud(CodRupt : String ; Axe : String ; ListeCodes : TStringList) : TSectRetri;
var
  Q : TQuery;
  Deb,Lon : TabByte;
  St,StTemp,Ru : String;
  i,Lg : Byte;
  Trie,Execute : Boolean;
begin
  Result := srOk;

  if Not VH^.Cpta[AxeToFb(Axe)].Structure then begin
    Result := srNonStruct;
    Exit;
  end;
  Ru := 'RU'+Axe[2];

  Q := OpenSQL('SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="'+Ru+'" AND CC_CODE="'+CodRupt+'"', True);
{$IFDEF EAGLCLIENT}
    if (Q.Fields[0].AsString = '') or (Q.Eof) then
{$ELSE}
    if (Q.Fields[0].AsString = '') or (IsFieldNull(Q,'CC_LIBRE')) or (Q.Eof) then
{$ENDIF}
    begin
    Result := srPasEnchainement;
    Exit;
    end
  else StTemp := Q.Fields[0].AsString;
  St:='' ;
  Ferme(Q);

  FillChar(Deb,SizeOF(Deb),#0);
  FillChar(Lon,SizeOF(Lon),#0);
  i:=1;
  While StTemp<>'' do begin
    St := ReadTokenSt(StTemp);
    Q := OpenSQL('SELECT SS_DEBUT,SS_LONGUEUR FROM STRUCRSE WHERE SS_AXE="'+Axe+'" AND SS_SOUSSECTION="'+st+'"', True);
    Deb[i] := Q.Fields[0].AsInteger;
    Lon[i] := Q.Fields[1].AsInteger;
    Ferme(Q);
    Inc(i);
  end;
  Lg := VH^.Cpta[AxeToFb(Axe)].Lg;

  Q := OpenSQL('SELECT BS_BUDSECT,BS_SECTIONTRIE FROM BUDSECT WHERE BS_AXE="'+Axe+'"', True);
  Trie := False;
  Execute := False;
  While (Not Q.EOF) AND (Not Trie) do begin
    St := '';
    for i:=1 to Lg do begin
      StTemp := Copy(Q.Fields[0].AsString,Deb[i],Lon[i]);
      if StTemp='' then Break;
      St := St + StTemp;
    end;

    if Not Execute then begin
{$IFDEF EAGLCLIENT}
        if (Q.Fields[1].AsString='') then
{$ELSE}
        if (Q.Fields[1].AsString='') or IsFieldNull(Q,'BS_SECTIONTRIE') then
{$ENDIF}
        begin
        Trie := False;
        Execute := True;
        end else
      if Q.Fields[0].AsString<>VH^.Cpta[AxeToFb(Axe)].Attente then begin
        Execute := True;
        if Q.Fields[1].AsString<>St then Trie:=False
                                    else Trie:=True;
      end;
    end;

    if Not Trie then begin
      ExecuteSQL('UPDATE BUDSECT SET BS_SECTIONTRIE="'+St+'" Where BS_BUDSECT="'+Q.Fields[0].AsString+'" And BS_AXE="'+Axe+'"');
    end;
    Q.Next;
  end;
  Ferme(Q);
end;

function EstMouvementeBudSect(const LeCpte, Laxe : String) : Boolean;
begin
  Result := ExisteSQL('SELECT BS_BUDSECT FROM BUDSECT WHERE BS_BUDSECT="'+LeCpte+'" '+
                      'AND BS_AXE="'+Laxe+'" And '+
                      '(EXISTS (SELECT BE_BUDSECT FROM BUDECR '+
                      'WHERE BE_BUDSECT="'+LeCpte+'" AND BE_AXE="'+Laxe+'"))');
end;

function ExisteBudSect(const LeCpte, Laxe : String) : Boolean;
begin
  Result := True;
  if (LeCpte='') or (Laxe='') then exit;
  if not ExisteSQL('SELECT BS_BUDSECT FROM BUDSECT WHERE BS_BUDSECT="'+LeCpte+'" AND BS_AXE="'+Laxe+'"') then begin
    MessageAlerte('L''enregistrement est inaccessible.');
    Result := False;
  end;
end;

procedure TOM_BUDSECT.OnNewRecord ;
begin
  Inherited ;

  UpdateDate;
  SetField('BS_SENS', 'M');
  if Mode in [taCreat..taCreatOne] then begin
    if AxeSerie<>'' then SetField('BS_AXE', AxeSerie)
                    else SetField('BS_AXE', 'A1');
    end
  else begin
    if Mode in [taModif] then begin
      if QuelAxe<>'' then SetField('BS_AXE', QuelAxe)
                     else SetField('BS_AXE', 'A1');
    end;
  end;


  SetControlChecked('BS_HT', False);
  SetControlEnabled('BS_BUDSECT', True);
  SetControlEnabled('BS_AXE', True);

  SetFocusControl('BS_BUDSECT');
  if Mode=taCreatEnSerie Then begin
    if SensDefaut<>'' then SetField('BS_SENS', SensDefaut);
    if SigneDefaut<>'' then SetField('BS_SIGNE', SigneDefaut);
    if AxeDefaut<>'' Then SetField('BS_Axe', AxeDefaut);
  end;
end ;

procedure TOM_BUDSECT.OnDeleteRecord ;
begin
  Inherited ;

  if EstMouvementeBudSect(GetField('BS_BUDSECT'), GetField('BS_AXE')) then begin
    LastError := 1;
    LastErrorMsg := MsgBox.Mess.Strings[7]; // Cette section budgétaire comporte des écritures budgétaires, vous ne pouvez pas la supprimer.
    Exit;
  end;
end;

procedure TOM_BUDSECT.OnUpdateRecord ;
begin
  Inherited ;

  if Not CodeValide then begin
    LastError :=1;
    exit;
  end;

  if DS.State in [dsEdit,dsInsert] then begin
   if GetField('BS_RUB') = '' then begin
     Msgbox.Execute(4,'',''); // Vous devez renseigner le code rubrique.
     CurseurAuCodeRub;
     LastError :=1;
     Exit;
   end;
   if FListe.Cells[0,1]='' then begin
     if MsgBox.Execute(5,'','')<>mrYes then begin // Vous n''avez pas associé de sections analytiques à votre section budgétaire. Désirez-vous continuer ?
       Fiche.Pages.ActivePage := Fiche.Pages.Pages[1];
       FListe.SetFocus;
       LastError :=1;
       Exit;
     end;
   end;
   if PresenceRub then begin
     if MsgBox.Execute(6,'','')<>mrYes then begin // Le code rubrique que vous avez renseigné existe déjà. La génération automatique des rubriques ne sera pas complète. Désirez-vous continuer ?
       CurseurAuCodeRub;
       LastError :=1;
       Exit;
     end;
   end;
  end;
  EcritCompteExclu;
end ;

procedure TOM_BUDSECT.OnAfterUpdateRecord ;
begin
  Inherited ;

  if mode = taCreatEnSerie then begin
    SensDefaut := GetField('BS_SENS');
    SigneDefaut := GetField('BS_SIGNE');
    AxeDefaut := GetField('BS_AXE');
  end;

  if mode in [taCreat..taCreatOne] then AxeSerie := GetField('BS_AXE');
end ;

procedure TOM_BUDSECT.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BUDSECT.OnLoadRecord ;
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
  MemoAxe := GetField('BS_AXE');
end ;

procedure TOM_BUDSECT.OnChangeField ( F: TField ) ;
begin
  Inherited ;

 {JP 11/08/05 : FQ 16045 : Déplacer dans CodeOnExit, car ici, c'est trop tôt
  if ((F.FieldName='BS_BUDSECT') and (GetField('BS_RUB')='')) then
    SetField('BS_RUB', Copy(F.AsString,1,5));}
end ;

// 1° : TActionFiche
// 2° : Axe
// 3° : Lequel ou Compte
// 4° : QuellePage
// 5° : LesModif
procedure TOM_BUDSECT.OnArgument ( S: String ) ;
var
  shtLibre : TTabSheet;
begin
  Inherited ;

  // Paramètres
  Mode := StrToTA(ReadTokenST(S));
  QuelAxe := ReadTokenST(S);
  Lequel := ReadTokenST(S);
  LaPage := READTOKENI(S);
  LesModif := ReadTokenST(S);

  MsgBox := THMsgBox.Create(Ecran);
  MsgBox.Mess.Add('0;Sections budgétaires;Le code que vous avez saisi existe déjà. Vous devez le modifier.;W;O;O;O;');
  MsgBox.Mess.Add('1;Sections budgétaires;Cette section budgétaire comporte des écritures budgétaires, vous ne pouvez pas modifier l''axe.;W;O;O;O;');
  MsgBox.Mess.Add('L''accès à ce compte  est interdit.');
  MsgBox.Mess.Add('3;Sections budgétaires;Votre saisie des comptes est trop longue, elle sera tronquée.;E;O;O;O;');
  MsgBox.Mess.Add('4;Sections budgétaires;Vous devez renseigner le code rubrique.;W;O;O;O;');
  MsgBox.Mess.Add('5;Sections budgétaires;Vous n''avez pas associé de sections analytiques à votre section budgétaire. Désirez-vous continuer ?;Q;YNC;N;N;');
  MsgBox.Mess.Add('6;Sections budgétaires;Le code rubrique que vous avez renseigné existe déjà. La génération automatique des rubriques ne sera pas complète. Désirez-vous continuer ?;Q;YNC;N;N;');
  MsgBox.Mess.Add('Cette section budgétaire comporte des écritures budgétaires, vous ne pouvez pas la supprimer.');

  shtLibre := TTabSheet(GetControl('ZL', True));
  if (shtLibre <> nil) then LibellesTableLibre(shtLibre, 'TBS_TABLE', 'BS_TABLE', 'D');

  Fiche := TFFiche(Ecran);
  FListe := THGrid(GetControl('FLISTE', True));
  FListe.OnKeyPress := FListeKeyPress;
  FListe.OnKeyDown := FListeKeyDown;
  FListe.OnCellExit := FListeCellExit;
  Fiche.BImprimer.Visible := True;
  Fiche.BImprimer.OnClick := BImprimerClick;
  THValComboBox(GetControl('BS_AXE', True)).OnChange := BS_AXEChange;

  {JP 11/08/05 : FQ 16045 : Pour renseigner le code rubrique par défaut}
  THEdit(GetControl('BS_BUDSECT', True)).OnExit := CodeOnExit;

  AxeSerie := '';
  MemoAxe := '';
  if Not OkConfidentiel then Exit;
  Fiche.Pages.ActivePage := Fiche.Pages.Pages[LaPage];
  case Mode Of
    taConsult : begin
      FListe.Enabled := True;
      FListe.Options := FListe.Options-[goEditing]+[goRowSelect];
    end;
    taCreat..taCreatOne : begin
      SetField('BS_BUDSECT', Lequel);
      Fiche.bDefaire.Enabled:=False;
    end;
  end;
  SensDefaut := '';
  SigneDefaut := '';
  AxeDefaut := 'A1';
  Ecran.HelpContext := 15133100;
end;

procedure TOM_BUDSECT.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_BUDSECT.OnCancelRecord ;
begin
  Inherited ;
end ;

function TOM_BUDSECT.Grid2St(ACol: Integer): String;
var
  i : Integer;
  St : String;
begin
  St:='';
  for i:=1 to FListe.RowCount-1 do
    St := St + FListe.Cells[ACol,i]+';';
  Result := St;
  if Length(Result)>250 then MsgBox.Execute(3,'',''); // Votre saisie des comptes est trop longue, elle sera tronquée.
end;

procedure TOM_BUDSECT.LitCompteExclu(St: String; ACol: Integer);
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

procedure TOM_BUDSECT.EcritCompteExclu;
begin
  SetField('BS_SECTIONRUB', Grid2St(0));
  SetField('BS_EXCLURUB', Grid2St(1));
end;

procedure TOM_BUDSECT.RempliGrid;
begin
  FListe.VidePile(False);
  LitCompteExclu(GetField('BS_SECTIONRUB'),0);
  LitCompteExclu(GetField('BS_EXCLURUB'),1);
end;

procedure TOM_BUDSECT.AffecteLe(Champ, Valeur: string);
var
  C : TControl;
begin
  C := GetControl(Champ);

  {$IFDEF EAGLCLIENT}
  if (C is TCheckBox) or (C is THValComboBox) or (C is TEdit) or (C is THEdit) then
  {$ELSE}
  if (C is TDBCheckBox) or (C is THDBValComboBox) or (C is TDBEdit) or (C is THDBCpteEdit) then
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

procedure TOM_BUDSECT.InitModifEnSerie(StModif: string);
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

function TOM_BUDSECT.CodeValide: Boolean;
var
  fb : TFichierBase;
begin
  Result := False;
  if DS.state in [dsInsert] then begin
    fb := TFichierBase(Ord(AxeToFb(GetField('BS_AXE')))-Ord(fbAxe1)+Ord(fbBudSec1));
    if ExisteCpte(GetField('BS_BUDSECT'),fb) then begin
      CurseurAuCode;
      MsgBox.Execute(0,'',''); // Le code que vous avez saisi existe déjà. Vous devez le modifier.
      Exit;
    end;
  end;
  Result := True;
end;

procedure TOM_BUDSECT.CurseurAuCode;
begin
  Fiche.Pages.ActivePage := Fiche.Pages.Pages[0];
  SetFocusControl('BS_BUDSECT');
end;

procedure TOM_BUDSECT.CurseurAuCodeRub;
begin
  Fiche.Pages.ActivePage := Fiche.Pages.Pages[0];
  SetFocusControl('BS_RUB');
end;

function TOM_BUDSECT.PresenceRub: Boolean;
begin
  if Length(GetField('BS_RUB'))>5 then SetField('BS_RUB', Copy(GetField('BS_RUB'),1,5));
  Result := PresenceComplexe('BUDSECT',['BS_BUDSECT','BS_RUB'],['<>','='],[GetField('BS_BUDSECT'),GetField('BS_RUB')],['S','S']);
end;

procedure TOM_BUDSECT.BImprimerClick(Sender: TObject);
begin
{$IFNDEF IMP}
PlanBudSec(GetField('BS_AXE'), GetField('BS_BUDSECT'), True);
{$ENDIF}
end;

procedure TOM_BUDSECT.FListeKeyPress(Sender: TObject; var Key: Char);
begin
  if Mode=taConsult then Exit;
  if Not (DS.State in [dsEdit,dsInsert]) then begin
    DS.Edit;
  end;
  Key := UpCase(Key);
end;

procedure TOM_BUDSECT.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Mode = taConsult then Exit;
  if Not(ssCtrl in Shift) then Exit;
  if Key<>VK_DELETE then Exit;
  if FListe.RowCount<=2 then Exit;

  FListe.DeleteRow(FListe.Row);
  if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount:=FListe.RowCount+1;
  if not(DS.Modified) then begin
    if not(DS.State in [dsEdit,dsInsert]) then DS.Edit;
  end;
end;

procedure TOM_BUDSECT.FListeCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if Mode = taConsult then Exit;
  if DS.Modified then
    if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.RowCount := FListe.RowCount+1;
end;

procedure TOM_BUDSECT.BS_AXEChange(Sender: TObject);
begin
  if DS.State<>dsEdit then Exit;
  if MemoAxe <> GetField('BS_AXE') then begin
    if EstMouvementeBudSect(GetField('BS_BUDSECT'), GetField('BS_AXE')) then begin
      MsgBox.Execute(1,'',''); // Cette section budgétaire comporte des écritures budgétaires, vous ne pouvez pas modifier l''axe.
      SetField('BS_AXE', MemoAxe);
    end;
  end;
end;

function TOM_BUDSECT.OkConfidentiel: Boolean;
begin
  Result:=False ;
  if V_PGI.Confidentiel='0' then begin
    if (GetCheckBoxState('BS_CONFIDENTIEL') = cbChecked) then begin
      MessageAlerte(MsgBox.Mess[2]); // L''accès à ce compte  est interdit.
      Close;
      Exit;
    end;
  end;
  SetControlVisible('BS_CONFIDENTIEL', (V_PGI.Confidentiel = '1'));
  SetControlEnabled('BS_CONFIDENTIEL', V_PGI.Superviseur);
  Result := True;
end;

procedure TOM_BUDSECT.UpdateDate;
var
  szDate : String;
begin
  szDate := FormatDateTime('dd mmm yyyy', GetField('BS_DATECREATION'));
  SetControlText('DATECREATION', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BS_DATEMODIF'));
  SetControlText('DATEMODIF', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BS_DATEOUVERTURE'));
  SetControlText('DATEOUVERTURE', szDate);
  szDate := FormatDateTime('dd mmm yyyy', GetField('BS_DATEFERMETURE'));
  SetControlText('DATEFERMETURE', szDate);
end;

{JP 11/08/05 : FQ 16045 : Pour renseigner le code rubrique par défaut
{---------------------------------------------------------------------------------------}
procedure TOM_BUDSECT.CodeOnExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if (GetField('BS_RUB') = '') and (DS.State in [dsInsert, dsEdit]) then
    SetField('BS_RUB', Copy(GetField('BS_BUDSECT'), 1, 5));
end;

Initialization
  registerclasses ( [ TOM_BUDSECT ] ) ;
end.
