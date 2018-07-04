{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 16/11/2001
Modifié le ... : 04/03/2003
Description .. : Source TOM de la TABLE : CPPRORATA (CPPRORATA)
Suite ........ : Passage en eAGL
Mots clefs ... : TOM;PRORATA
*****************************************************************}
// Lors de la remise en place du champ PA_TOUSJAL ,
// décommenté les lignes avec *+* et supprimer les lignes avec *-*
Unit CPPRORATA_TOM ;

Interface

Uses SysUtils,       // StrToFloat
     Classes,        // registerclasses
{$IFDEF EAGLCLIENT}
     MaineAGL,       // AGLLanceFiche
     StdCtrls,       // TCheckbox
{$ELSE}
     Fe_main,        // AGLLanceFiche
     DB,             // TField
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     DBCTRLS,        // TDBCheckbox
{$ENDIF}
     HCtrls,         // ExisteSQL, OpensQL, Ferme, THValComboBox
     HMsgBox,        // PGIBox
     Utom,           // TOM
     Utob,           // TOB
     HEnt1           // TraduireMemoire
     ;

Const
 titre = 'Prorata de TVA' ;

procedure CCLanceFiche_ParamTauxTVA;

Type
  TOM_CPPRORATA = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private
    procedure ExerciceOnChange(Sender: TObject);
    procedure MajFormat(VF: Boolean);
    function AvecMvt1: Boolean;
    function AvecMvt2: Boolean;
    function ExisteEcr: Boolean;
    function AvecMvt3: Boolean;
// *+*     procedure TousJalClick(Sender : TObject);
    end ;

Implementation

var Tx1, Tx2, Tx3 : String ;

procedure CCLanceFiche_ParamTauxTVA;
begin
  AGLLanceFiche ('CP' , 'CPTABPRORATA', '', '' , '' ) ;
end;

procedure TOM_CPPRORATA.OnNewRecord ;
begin
  Inherited ;
// *+*   if bTousJal then SetControlEnabled('PA_TOUSJAL', True);
// *+*   THValComboBox(GetControl('PA_JOURNAL', True)).ItemIndex := -1;
end ;

procedure TOM_CPPRORATA.OnDeleteRecord ;
begin
  SetControlvisible('ALERTE',True) ;
  if ExisteEcr then begin
    Lasterror:=-1 ;
    SetControlvisible('ALERTE',False) ;
    PgiBox(TraduireMemoire('Vous ne pouvez détruire cet enregistrement. Des écritures ont été modifiées avec ces taux.'),TraduireMemoire('Prorata de TVA')) ;
    exit ;
  end ;
  SetControlvisible('ALERTE',False) ;
  Inherited ;
end ;

Function TOM_CPPRORATA.ExisteEcr : Boolean ;
Begin
  Result:=False ;

// *+* if (GetControlText('PA_TOUSJAL') <> 'X') then begin
  if (GetControlText('PA_JOURNAL') <> '') then begin // *-*
    if ExisteSql('SELECT E_JOURNAL FROM ECRITURE WHERE (E_QUALIFORIGINE="TV2" OR E_QUALIFORIGINE="TV1" OR E_QUALIFORIGINE="TVI") AND E_EXERCICE="'+GetControlText('PA_CODE')+'" AND E_JOURNAL="'+GetControlText('PA_JOURNAL')+'"') then Result:=True;
  end;
End ;

Procedure TOM_CPPRORATA.MajFormat (VF : Boolean ) ;
Begin
  If VF then begin
    SetControlProperty('PA_TAUX1','DisplayFormat','##0.##%') ;
    SetControlProperty('PA_TAUX2','DisplayFormat','##0.##%') ;
    SetControlProperty('PA_TAUX3','DisplayFormat','##0.##%') ;
  end
  else begin
    SetControlProperty('PA_TAUX1','DisplayFormat','##0.##') ;
    SetControlProperty('PA_TAUX2','DisplayFormat','##0.##') ;
    SetControlProperty('PA_TAUX3','DisplayFormat','##0.##') ;
  end ;
End ;

procedure TOM_CPPRORATA.OnUpdateRecord ;
var
  sz1,sz2,sz3 : String;
  Requete     : string;
begin
Inherited ;
LastErrorMsg := '';
{JP 06/06/2003 : on n'autorise pas de créer des couples Exercice/Journal avec le couple Exercice/<<Tous>>}
// *+*  if (GetControlText('PA_TOUSJAL') = 'X') then
// *+*    Requete := 'SELECT PA_JOURNAL FROM CPPRORATA WHERE PA_CODE = "' + GetControlText('PA_CODE') + '"'
  Requete := 'SELECT PA_JOURNAL FROM CPPRORATA WHERE PA_CODE = "' + GetControlText('PA_CODE') + '" AND ('; // *-*
  if GetControlText('PA_JOURNAL') <> '' then // *-*
    Requete := Requete + ' PA_JOURNAL = "---") '// *-*
  else
    Requete := Requete + ' PA_JOURNAL <> "---")'; // *-*
// *+*    Requete := 'SELECT PA_JOURNAL FROM CPPRORATA WHERE PA_CODE = "' + GetControlText('PA_CODE') + '" AND PA_TOUSJAL="X"';
  if ExisteSql(Requete) then begin
    LastError := 1;
    LastErrorMsg := 'Il n''est pas possible pour un même exercice, de paramétrer'#13 +
                    'différents taux de prorata de TVA pour un même journal !';
    Exit;
  end;
SetControlvisible('ALERTE',True) ;
if (AvecMvt1)  then
  begin
  Lasterror := -1 ;
  PgiBox(TraduireMemoire('Des écritures sont déjà "proratisées" avec ce taux. Vous ne pouvez pas le modifier.'),TraduireMemoire(Titre)) ;
  MajFormat(True) ;
  SetFocusControl('PA_TAUX1') ;
  SetControlvisible('ALERTE',False) ;
  Exit ;
  end ;
if (AvecMvt2) then
  begin
  Lasterror := -1 ;
  PgiBox(TraduireMemoire('Des écritures sont déjà "proratisées" avec ce taux. Vous ne pouvez pas le modifier.'),TraduireMemoire(Titre)) ;
  MajFormat(True) ;
  SetFocusControl('PA_TAUX2') ;
  SetControlvisible('ALERTE',False) ; Exit ;
  end ;
if (AvecMvt3) then
  begin
  Lasterror := -1 ;
  PgiBox(TraduireMemoire('Des écritures sont déjà "proratisées" avec ce taux. Vous ne pouvez pas le modifier.'),TraduireMemoire(Titre)) ;
  MajFormat(True) ;
  SetFocusControl('PA_TAUX3') ;
  SetControlvisible('ALERTE',False) ; Exit ;
  end ;
SetControlvisible('ALERTE',False) ;
MajFormat(False) ;

  sz1 := GetControlText('PA_TAUX1');
  if (Pos('%',sz1)>0) then Delete(sz1,Length(sz1),1);
  If (StrToFloat(sz1) > 100) or (StrToFloat(sz1) < 0) then begin
    LastError := -1 ;
    PgiBox(TraduireMemoire('Le prorata doit être compris entre 0 et 100 %'),TraduireMemoire(Titre)) ;
    MajFormat(True) ;
    SetFocusControl('PA_TAUX1') ;
    Exit ;
  end ;

  sz2 := GetControlText('PA_TAUX2');
  if (Pos('%',sz2)>0) then Delete(sz2,Length(sz2),1);
  If (StrToFloat(sz2) > 100) or (StrToFloat(sz2) < 0) then begin
    LastError := -1 ;
    PgiBox(TraduireMemoire('Le prorata doit être compris entre 0 et 100 %'),TraduireMemoire(Titre)) ;
    MajFormat(True) ;
    SetFocusControl('PA_TAUX2') ;
    Exit ;
  end ;

  sz3 := GetControlText('PA_TAUX3');
  if (Pos('%',sz3)>0) then Delete(sz3,Length(sz3),1);
  If (StrToFloat(sz3) > 100) or (StrToFloat(sz3) < 0) then begin
    LastError := -1 ;
    PgiBox(TraduireMemoire('Le prorata doit être compris entre 0 et 100 %'),TraduireMemoire(Titre)) ;
    MajFormat(True) ;
    SetFocusControl('PA_TAUX3') ;
    Exit ;
  end ;

  { // FQ 12252 Point 1
  if (StrToFloat(sz2)> 0) and
    ((StrToFloat(sz1)= 0) or
     (StrToFloat(sz3)= 0)) then begin
    LastError := -1 ;
    PgiBox(TraduireMemoire('Si le taux définitif est renseigné, les autres taux doivent l''être aussi.'),TraduireMemoire(Titre)) ;
    MajFormat(True) ;
    SetFocusControl('PA_TAUX1');
    Exit ;
  end ;

  If (StrToFloat(sz3) = StrToFloat(sz2)) then begin
    LastError := -1 ;
    PgiBox(TraduireMemoire('Le taux intermédiaire et définitif ne peuvent pas être égaux.'),TraduireMemoire(Titre)) ;
    MajFormat(True) ;
    SetFocusControl('PA_TAUX3') ;
    Exit ;
  end ;}

  // FQ 12252 Point 3
  if (StrToFloat(sz2)= 0) and
    ((StrToFloat(sz1)= 0) and
     (StrToFloat(sz3)= 0)) then begin
    LastError := -1 ;
    PgiBox(TraduireMemoire('Vous ne pouvez pas laisser tous les taux à 0.'),TraduireMemoire(Titre)) ;
    MajFormat(True) ;
    SetFocusControl('PA_TAUX1') ;
    Exit ;
  end ;

  if (GetControlText('PA_JOURNAL') = '') then // *-*
    SetField('PA_JOURNAL', '---'); // *-*

  MajFormat(True) ;

// *+*  SetControlEnabled('PA_TOUSJAL', False);
end ;

procedure TOM_CPPRORATA.OnAfterUpdateRecord ;
begin
  Inherited ;
  // 13269
  AvertirCacheServer('CPPRORATA');
end ;

procedure TOM_CPPRORATA.OnLoadRecord ;
begin
  Tx1 := GetControlText('PA_TAUX1') ;
  Tx2 := GetControlText('PA_TAUX2') ;
  Tx3 := GetControlText('PA_TAUX3') ;
  Inherited ;
end ;

Function TOM_CPPRORATA.AvecMvt1 :Boolean;
begin
Result:=False ;
  if GetControlText('PA_TAUX1') <> Tx1 then begin
//*+*    if (GetControlText('PA_TOUSJAL') <> 'X') then begin
    if (GetControlText('PA_JOURNAL') <> '') then begin // *-*
      if ExisteSql('SELECT E_JOURNAL FROM ECRITURE WHERE E_QUALIFORIGINE="TV1" AND E_EXERCICE="'+GetControlText('PA_CODE')+'" AND E_JOURNAL="'+GetControlText('PA_JOURNAL')+'"') then Result:=True;
      end
    else
      if ExisteSql('SELECT E_JOURNAL FROM ECRITURE WHERE E_QUALIFORIGINE="TV1" AND E_EXERCICE="'+GetControlText('PA_CODE')+'"') then Result:=True ;
  end;
end ;

Function TOM_CPPRORATA.AvecMvt2 : Boolean ;
begin
Result:=False ;
  If GetControlText('PA_TAUX2') <> Tx2 then begin
// *+*    if (GetControlText('PA_TOUSJAL') <> 'X') then begin
    if (GetControlText('PA_JOURNAL') <> '') then begin // *-*
      if ExisteSql('SELECT E_JOURNAL FROM ECRITURE WHERE E_QUALIFORIGINE="TV2" AND E_EXERCICE="'+GetControlText('PA_CODE')+'" AND E_JOURNAL="'+GetControlText('PA_JOURNAL')+'"') then Result:=True;
      end
    else
      if ExisteSql('SELECT E_JOURNAL FROM ECRITURE WHERE E_QUALIFORIGINE="TV2" AND E_EXERCICE="'+GetControlText('PA_CODE')+'"') then Result:=True ;
  end;
end ;

Function TOM_CPPRORATA.AvecMvt3 : Boolean ;
begin
Result:=False ;
  If GetControlText('PA_TAUX3') <> Tx3 then begin
// *+*    if (GetControlText('PA_TOUSJAL') <> 'X') then begin
    if (GetControlText('PA_JOURNAL') <> '') then begin // *-*
      if ExisteSql('SELECT E_JOURNAL FROM ECRITURE WHERE E_QUALIFORIGINE="TVI" AND E_EXERCICE="'+GetControlText('PA_CODE')+'" AND E_JOURNAL="'+GetControlText('PA_JOURNAL')+'"') then Result:=True;
      end
    else
      if ExisteSql('SELECT E_JOURNAL FROM ECRITURE WHERE E_QUALIFORIGINE="TVI" AND E_EXERCICE="'+GetControlText('PA_CODE')+'"') then Result:=True ;
  end;
end ;

procedure TOM_CPPRORATA.OnChangeField ( F: TField ) ;
var
  i : Integer;
  Journal : THValComboBox;
begin
 Inherited ;
 if (F.FieldName = 'PA_JOURNAL') and (F.AsString = '---') then // *-*
   THValComboBox(GetControl('PA_JOURNAL')).ItemIndex := 0; // *-*

 // FQ 12913
 if f.FieldName = 'PA_JOURNAL' then begin
   Journal := THValComboBox(GetControl('PA_JOURNAL'));
   for i := 0 to Journal.Items.Count-1 do begin
     if (Journal.Values.Strings[i] = F.AsString) then begin
       Journal.itemindex := i;
       exit;
     end;
   end;
 end;
end ;


Procedure TOM_CPPRORATA.ExerciceOnChange (Sender : TObject) ;
var
  Q : TQuery ;
    TobDat : Tob ;
BEGIN
  if ((GetControlText('PA_CODE')='') or (DS.State <> dsInsert)) then Exit;
TobDat := Tob.Create('Dat', nil, -1) ;
Q := OpenSql('SELECT EX_DATEDEBUT, EX_DATEFIN, EX_LIBELLE FROM EXERCICE WHERE EX_EXERCICE="'+GetControlText('Pa_Code')+'"', True) ;
TobDat.LoadDetailDb('Dat','','',Q,False,True) ;
Ferme(Q) ;
  SetControlEnabled('PA_DATEDEBUT', True) ;
  SetControlEnabled('PA_DATEINTER', True) ;
  SetFocusControl('PA_DATEDEBUT') ;  SetControlText('PA_DATEDEBUT', TobDat.Detail[0].GetValue('Ex_DATEDEBUT')) ;//SetField('PA_DATEDEBUT',DateToStr(Q.Fields[0].AsDateTime)) ;
  SetFocusControl('PA_DATEINTER') ;  SetControlText('PA_DATEINTER', TobDat.Detail[0].GetValue('EX_DATEFIN')) ;//SetField('PA_DATEINTER',DateToStr(Q.Fields[1].AsDateTime)) ;
  SetFocusControl('PA_DATEINTER2') ; SetControlText('PA_DATEINTER2',TobDat.Detail[0].GetValue('EX_DATEDEBUT')) ;//
  SetControlEnabled('PA_DATEDEBUT', False) ;
  SetControlEnabled('PA_DATEINTER', False) ;
SetFocusControl('PA_LIBELLE');
SetControlText('PA_LIBELLE', TobDat.Detail[0].GetValue('EX_LIBELLE')) ; //Libelle.SetFocus ; Libelle.Text := Q.Fields[2].AsString  ;
TobDat.Free ; //TobDat := Nil ;
END ;


procedure TOM_CPPRORATA.OnArgument ( S: String ) ;
begin
Ecran.HelpContext := 1546;
THValComboBox(GetControl('PA_CODE')).OnChange := ExerciceOnChange ;
// *+* TDBCheckbox(GetControl('PA_TOUSJAL', True)).OnClick := TousJalClick;
Inherited ;
end ;

procedure TOM_CPPRORATA.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CPPRORATA.OnCancelRecord ;
begin
  Inherited ;
end ;

(* *+* procedure TOM_CPPRORATA.TousJalClick(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('PA_JOURNAL', True)).Enabled := not TCheckbox(GetControl('PA_TOUSJAL', True)).Checked;
{$ELSE}
  THDBValComboBox(GetControl('PA_JOURNAL', True)).Enabled := not TDBCheckbox(GetControl('PA_TOUSJAL', True)).Checked;
{$ENDIF}
end; *)

Initialization
  registerclasses ( [ TOM_CPPRORATA ] ) ;
end.
