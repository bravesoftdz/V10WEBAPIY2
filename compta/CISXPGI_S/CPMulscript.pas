unit CPMulscript;

interface

uses
  Classes, Controls,
{$IFDEF EAGLCLIENT}
     MaineAGL, UtileAgl, UTob, eMul, UScriptTob,
{$ELSE}
     db, Mul, HDB,
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
     FE_main,
{$ENDIF}

  StdCtrls, Hctrls, 
  hmsgbox, HEnt1, UControlParam,
  UTOF, UPDomaine, UAssistConcept, UScript, HSysMenu;


Type
  TOF_MULSCRIPT = Class (TOF)
       procedure FListeDblClick(Sender: TObject);
       procedure OnArgument (S : String ) ; override ;
       procedure InsertClick(Sender: TObject);
       procedure DeleteClick(Sender: TObject);
       private
       ModeSelect   : Boolean;
       {$IFDEF EAGLCLIENT}
       FListe : THGrid;
       {$ELSE}
       FListe : THDBGrid;
       {$ENDIF}

  end ;

implementation

uses UConcept;

procedure TOF_MULSCRIPT.FListeDblClick(Sender: TObject);
var
  Table, Dm, Nat: string;
  lib,Nt,edt    : string;
  Delim         : Boolean;
  TPControle    : TSVControle;
begin
{$IFDEF EAGLCLIENT} // positionnement de Query
 TFMul(Ecran).Q.TQ.Seek(THGrid(TFMul(Ecran).FListe).Row - 1);
{$ENDIF}
  Table := TFMul(Ecran).Q.FindField('CIS_CLE').AsString;
  Dm    := TFMul(Ecran).Q.FindField('CIS_Domaine').AsString;
  if ModeSelect then
  begin
       TFMul(Ecran).ModalResult := MrOk;
       TFMul(Ecran).retour:= Table + ';'+ Dm;
       exit;
  end;
  Nat   := TFMul(Ecran).Q.FindField('CIS_CLEPAR').AsString;
  lib   := RendLibelleDomaine(Dm);
  Nt    :=  TFMul(Ecran).Q.FindField('CIS_Table2').AsString;
  Edt   := TFMul(Ecran).Q.FindField('CIS_Table1').AsString;
  Delim := (TFMul(Ecran).Q.FindField('CIS_Nature').AsString = 'X');
  if Delim then
     ModifScriptDelim(Table, Nat, Nt, Edt, Dm)
  else
  begin
      TPControle := TSVControle.create;
      TPControle.LePays := TFMul(Ecran).Q.FindField('CIS_Table3').AsString;
      TPControle.ChargeTobParam(Dm);
      PExtConcept(nil, Lib, Table, Nat, Nt, Edt, taModif, TPControle);
  end;
end;

procedure TOF_MULSCRIPT.InsertClick(Sender: TObject);
begin
  CreateScript (FALSE, tacreat, '', TRUE, THEdit(GetControl('CIS_CLEPAR')).Text, THValComboBox(GetControl('CIS_TABLE3')).value);
  TFMul(Ecran).FListe.refresh;
end;

procedure TOF_MULSCRIPT.DeleteClick(Sender: TObject);
begin
  // Supppression
  if executeSql('DELETE FROM CPGZIMPREQ WHERE CIS_CLE="' + GetField('CIS_CLE') + '"') <> 1
    then PGIInfo('Le script selectionné n''a pas pu être supprimé.', Ecran.Caption);
 // On relnce la recherche
  TFMul(Ecran).BChercheClick(nil);
end;


procedure TOF_MULSCRIPT.OnArgument (S : String );
var
St : string;
begin
TFMul(Ecran).FListe.OnDblClick := FListeDblClick;
TFMul(Ecran).BInsert.OnClick := InsertClick;
TButton(GetControl('BDelete',True)).onClick  := DeleteClick;

ModeSelect := (S ='TRUE');
St := ReadTokenPipe(S, '=');
if S <> '' then
begin
         St := ReadTokenSt(S);
         if St <>'' then
         THEdit(GetControl('CIS_CLEPAR')).Text := St;
         St := ReadTokenPipe (S, '=');
         if S <> '' then
            THValComboBox(GetControl('CIS_TABLE3')).Itemindex := THValComboBox(GetControl('CIS_TABLE3')).Values.IndexOf(S);

         if not V_PGI.SAV then
         begin
              SetControlEnabled ('CIS_CLEPAR', FALSE);
              if (GetControlText('CIS_CLEPAR') = 'EXPORT') and
              (GetControlText('CIS_TABLE3') <> '' ) then SetControlEnabled ('CIS_TABLE3', FALSE);
         end;
end;
if (GetInfoVHCX.Domaine <> '') then
begin
         THValComboBox(GetControl('CIS_DOMAINE')).value := GetInfoVHCX.Domaine;
{$IFDEF COMPTA}
         SetControlEnabled ('CIS_DOMAINE', FALSE);
{$ENDIF}

{$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FListe'));
    TFMul(Ecran).HMTrad.ResizeGridColumns(FListe);
{$ELSE}
    FListe := THDBGrid (GetControl('FListe'));
    CentreDBGrid(FListe);
    THSystemMenu(GetControl('HMTrad')).ResizeDBGridColumns(FListe);
{$ENDIF}


end;


end;



initialization
  registerclasses([TOF_MULSCRIPT]);

end.
