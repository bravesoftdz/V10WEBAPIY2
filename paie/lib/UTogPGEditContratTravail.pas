{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/02/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGEDITCONTRATTRAVAIL ()
Mots clefs ... : TOF;PGEDITCONTRATTRAVAIL
*****************************************************************
PT 23/11/2005 JL V_65 FQ 12514 Gestion des champs XX_WHERE et XX_ORDERBY
}
Unit UTogPGEditContratTravail ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HQRy,
     EntPaie,
     P5DEF,
     ParamDat,
     PGOutils2,
     UTOF ;

Type
  TOF_PGEDITCONTRATTRAVAIL = Class (TOF)
     procedure OnUpdate                 ; override ;
     procedure OnArgument (S : String ) ; override ;
    private
    procedure ExitEdit(Sender: TObject);
    procedure CkSortieClick (Sender : TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure ChangeCheckBox(Sender : TObject);
  end ;

Implementation

procedure TOF_PGEDITCONTRATTRAVAIL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGEDITCONTRATTRAVAIL.OnArgument (S : String ) ;
var Num : Integer;
    Edit : THEdit;
    Check : TCheckBox;
begin
  Inherited ;
  SetControlText('PSA_DATEENTREE',DateToStr(IDate1900));
  SetControlText('PSA_DATEENTREE_',DateToStr(V_PGI.DateEntree));
  For Num  :=  1 to VH_Paie.PGNbreStatOrg do
  begin
       if Num >4 then Break;
       VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
       VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)+'_'),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)+'_'));
       SetControlVisible('CN'+IntToStr(Num),True);
  end;
  VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ;
  SetControlVisible('CN'+IntToStr(Num),(VH_Paie.PGLibCodeStat <> ''));
  Edit := THEdit(GetControl('PSA_SALARIE'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
  Edit := THEdit(GetControl('PSA_SALARIE_'));
  if Edit <> nil then Edit.OnExit := ExitEdit;
  Check := TCheckBox(GetControl('CKSORTIE'));
  If Check <> Nil Then Check.OnCLick := CkSortieClick;
  Edit := THEdit(GetControl('DATESORTIE'));
  If Edit <> Nil Then Edit.OnElipsisClick := DateElipsisclick;
  Edit := THEdit(GetControl('PSA_DATEENTREE'));
  If Edit <> Nil Then Edit.OnElipsisClick := DateElipsisclick;
  Edit := THEdit(GetControl('PSA_DATEENTREE_'));
  If Edit <> Nil Then Edit.OnElipsisClick := DateElipsisclick;
  Check := TCheckBox(GetControl('CETAB'));
  If Check <> Nil Then Check.OnCLick := ChangeCheckBox;
  Check := TCheckBox(GetControl('CALPHA'));
  If Check <> Nil Then Check.OnCLick := ChangeCheckBox;
end ;

procedure TOF_PGEDITCONTRATTRAVAIL.ExitEdit(Sender: TObject);
var
  edit: thedit;
begin
  edit := THEdit(Sender);
  if edit <> nil then
    if (VH_Paie.PgTypeNumSal = 'NUM') and (length(Edit.text) < 11) and (isnumeric(edit.text)) then
      edit.text := AffectDefautCode(edit, 10);
end;

procedure TOF_PGEDITCONTRATTRAVAIL.CkSortieClick(Sender : TObject);
var STWhere : String;
begin
        If Sender = Nil then Exit;
        If GetCheckBoxState('CKSORTIE') = CbChecked then
        begin
                SetControlVisible('DATESORTIE',True);
                SetControlText('DATESORTIE',DateToStr(Date));
                SetControlVisible('TDATEARRET',True);
                StWhere := '(PSA_DATESORTIE<="'+UsDateTime(IDate1900)+'" '+
                'OR PSA_DATESORTIE IS NULL OR PSA_DATESORTIE>="'+UsDateTime(StrToDate(GetControlText('DATESORTIE')))+'")';
        end
        else
        begin
                SetControlVisible('DATESORTIE',False);
                SetControlVisible('TDATEARRET',False);
                StWhere := '';
        end;
        SetControlText('XX_WHERE',StWhere);
end;

procedure TOF_PGEDITCONTRATTRAVAIL.DateElipsisclick(Sender: TObject);
var key : char;
begin
    key := '*';
    ParamDate (Ecran, Sender, Key);
end;

procedure TOF_PGEDITCONTRATTRAVAIL.ChangeCheckBox(Sender : TObject);
var StOrder : String;
begin
  StOrder := '';
  If GetCheckBoxState('CETAB') = CbChecked then StOrder := 'PCI_ETABLISSEMENT';
  If GetCheckBoxState('CALPHA') = CbChecked then
  begin
       iF StOrder <> '' then StOrder := StOrder + ',PSA_LIBELLE'
       Else StOrder := ' PSA_LIBELLE';
  end;
  If StOrder <> '' then StOrder := StOrder + ',PCI_SALARIE'
  Else StOrder := 'PCI_SALARIE';
  SetControlText('XX_ORDERBY',StOrder);
end;


Initialization
  registerclasses ( [ TOF_PGEDITCONTRATTRAVAIL ] ) ;
end.

