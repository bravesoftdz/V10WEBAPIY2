{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 05/01/2005
Modifié le ... : 05/01/2005
Description .. : Le source 2/3 n'est pas testé
Mots clefs ... : TOF;MULMVTBU
*****************************************************************}
Unit MULMVTBU_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     eMul, uTob, MaineAgl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     db, mul, FE_Main,
{$ENDIF}
     eSaisBud, // SaisieBudget
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, SAISUTIL,
     Ent1,
     UTOF;

procedure MultiCritereMvtBud(Comment : TActionFiche ; QuelleSaisie : String);

Type
  TOF_MULMVTBU = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    HM : THMsgBox;
    Mul : TFMul;
    BE_EXERCICE : THValComboBox;
    BE_DATECOMPTABLE,BE_DATECOMPTABLE_ : THEdit;
    Comment : TActionFiche;
    procedure FListeDblClick(Sender: TObject);
    procedure BE_EXERCICEChange(Sender: TObject);
  end;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  AglInit;

procedure MultiCritereMvtBud(Comment : TActionFiche ; QuelleSaisie : String);
var
  M : RMVT;
begin
  if Comment=taCreat then begin
    FillChar(M,Sizeof(M),#0);
    M.Etabl:=VH^.EtablisDefaut;
    M.CodeD:=V_PGI.DevisePivot;
    M.Simul:='N';
    M.Nature:='INI';
    M.TypeSaisie:=QuelleSaisie;
    SaisieBudget(taCreat,M,False);
    Exit;
  end;
  AGLLanceFiche('CP', 'CPMULMVTBU', '', '', ActionToString(Comment)+';'+QuelleSaisie);
end;

procedure TOF_MULMVTBU.OnNew ;
begin
  Inherited ;
end;

procedure TOF_MULMVTBU.OnDelete ;
begin
  Inherited ;
end;

procedure TOF_MULMVTBU.OnUpdate ;
begin
  Inherited ;
end;

procedure TOF_MULMVTBU.OnLoad ;
begin
  Inherited ;
end;

procedure TOF_MULMVTBU.OnArgument (S : String ) ;
var
  sz : String;
  shtLibre : TTabSheet;
begin
  HM := THMsgBox.Create(Ecran);
  HM.Mess.Add('Visualisation des écritures bugétaires');
  HM.Mess.Add('Modification des écritures budgétaires');

  // FormShow
  sz := S;
  Comment := StringToAction(ReadTokenST(sz));
  Case Comment Of
    taConsult : begin
      Ecran.Caption:= HM.Mess[0];
      Ecran.HelpContext:=15221000;
    end;
    taModif   : begin
      Ecran.Caption:= HM.Mess[1];
      Ecran.HelpContext:=15230000; TWinControl(GetControl('PCritere',True)).HelpContext:=15230000;
    end;
  end;
  UpdateCaption(Ecran);

  shtLibre := TTabSheet(GetControl('PZLIBRE', True));
  if (shtLibre <> nil) then LibellesTableLibre(shtLibre,'TBE_TABLE','BE_TABLE','U');
  PositionneEtabUser(GetControl('BE_ETABLISSEMENT', True));

  Inherited;

  // Evènements
  BE_EXERCICE := THValComboBox(GetControl('BE_EXERCICE', True));
  BE_EXERCICE.OnChange := BE_EXERCICEChange;
  BE_DATECOMPTABLE := THEdit(GetControl('BE_DATECOMPTABLE', True));
  BE_DATECOMPTABLE_ := THEdit(GetControl('BE_DATECOMPTABLE_', True));
  Mul := TFMul(Ecran);
  Mul.FListe.OnDblClick := FListeDblClick;
end;

procedure TOF_MULMVTBU.OnClose ;
begin
  Inherited ;

  HM.Free;
end;

procedure TOF_MULMVTBU.OnDisplay () ;
begin
  Inherited ;
end;

procedure TOF_MULMVTBU.OnCancel () ;
begin
  Inherited ;
end;

procedure TOF_MULMVTBU.FListeDblClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  TOBListe, TOBLigne : TOB;
{$ENDIF}
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit ;

{$IFDEF EAGLCLIENT}
  TOBListe := TOB.Create('Liste Ecriture',nil,-1) ;
  TOBLigne := Mul.Q.TQ.Detail[Mul.FListe.row-1] ;
  TOB.Create('BUDECR',TOBListe,-1) ;
  TOBListe.Detail[0].Dupliquer(TOBLigne, False, True);
  TrouveEtLanceSaisBud(TOBListe, Comment);
  FreeAndNil(TOBListe);
{$ELSE}
  TrouveEtLanceSaisBud(Mul.Q, Comment);
{$ENDIF}

end;

procedure TOF_MULMVTBU.BE_EXERCICEChange(Sender: TObject);
begin
  ExoToDates(BE_EXERCICE.Value,BE_DATECOMPTABLE,BE_DATECOMPTABLE_);
end;

Initialization
  registerclasses ( [ TOF_MULMVTBU ] ) ; 
end.
