{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : RECUPDONNEES ()
Mots clefs ... : TOF;RECUPDONNEES
*****************************************************************}
Unit uTofRecupDonnees ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
     HCtrls, HEnt1, HMsgBox, UTOF, PGIEnv, IntegGen, IntegBase, IntegExec,Ent1 ,HPanel, Vierge;

Type
  TOF_RECUPDONNEES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  private
    fContexte : TIntegContexte;
    procedure ChargeListeBase;
    function GetBaseInfo(stFileName: string; var stBaseName: string;
      TypeBase: string): boolean;
    procedure ChargeContexte;
    procedure LibereContexte;
    procedure OnClickRB (Sender : TObject);
  end ;

Implementation

procedure TOF_RECUPDONNEES.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RECUPDONNEES.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RECUPDONNEES.OnUpdate ;
var Base : TIntegBase;
    Exec : TIntegExec;
    stFileDest : string;
begin
  Inherited ;
  ChargeContexte;
  Base := TIntegBase.Create(FContexte.Base);
  Base.Charge;
  Exec := TIntegExec.Create (Base,FContexte);
  Exec.BuildTOB;
  if GetControlText('ZMODEINTEGRATION')='DOS' then Exec.ToDossier
  else if GetControlText('ZMODEINTEGRATION')='PGI' then
  begin
    stFileDest := Copy (FContexte.FileName,1,Pos(FContexte.FileName,'.'))+'.PGI';
    Exec.ToFilePGI (stFileDest);
  end;
  Base.Free;
  LibereContexte;
end ;

procedure TOF_RECUPDONNEES.OnLoad ;
begin
  Inherited ;
  SetControlChecked ('ZRBGL',TRUE);
  SetControlChecked ('ZRBBAL',FALSE);
  ChargeListeBase;
end ;

procedure TOF_RECUPDONNEES.OnArgument (S : String ) ;
begin
  Inherited ;
  TRadioButton (GetControl('ZRBGL')).OnClick := OnClickRB;
  TRadioButton (GetControl('ZRBBAL')).OnClick := OnClickRB;
end ;

procedure TOF_RECUPDONNEES.OnClose ;
begin
inherited;
end ;

procedure TOF_RECUPDONNEES.ChargeListeBase;
var SearchRec: TSearchRec;
    r_search  : integer;
    stBaseName, stTypeBase : string;
    ZBase : THValComboBox;
begin
  ZBase := THValComboBox (GetControl('ZBase'));
  ZBase.Items.Clear;
  ZBase.Values.Clear;
  if GetControlText ('ZRBGL')='X' then stTypeBase := 'GL' else
  stTypeBase := 'BL';
  r_search := FindFirst('c:\pgi00\std\*.*', faAnyFile, SearchRec);
  while (r_search = 0) do
  begin
    if SearchRec.Attr <> faDirectory then
    begin
      if GetBaseInfo('c:\pgi00\std\'+SearchRec.Name, stBaseName,stTypeBase) then
      begin
        ZBase.Items.Add (stBaseName);
        ZBase.Values.Add ('c:\pgi00\std\'+SearchRec.Name);
      end;
    end;
    r_search := FindNext(SearchRec);
  end;
  FindClose(SearchRec);
end;

function TOF_RECUPDONNEES.GetBaseInfo (stFileName : string; var stBaseName : string; TypeBase: string) : boolean;
var F : TextFile;
  bTrouve : boolean;
  St : string;
begin
  bTrouve := False;
  AssignFile (F,stFileName);
  Reset (F);
  Readln (F, St);
  if IntegExtractEntete(St) = EX_TYPE then
  begin
    Readln (F, St);
    if IntegExtractEntete(St) = EX_TITRE then
    begin
      stBaseName := IntegExtractChaine (St);
      while not Eof (F) do
      begin
        Readln (F, St);
        if IntegExtractEntete(St) = EX_TYPERECUP then
        begin
          if IntegExtractChaine (St)=TypeBase then
          begin
            bTrouve := True;
            break;
          end;
        end;
      end;
    end;
  end;
  CloseFile (F);
  Result := bTrouve;
end;

procedure TOF_RECUPDONNEES.OnClickRB(Sender: TObject);
begin
  inherited;
  ChargeListeBase;
end;

procedure TOF_RECUPDONNEES.ChargeContexte;
var i,j : integer;
  PColl : PIntegCollectif;
  PTabConv : PIntegTabConv;
  bOk : boolean;
  Q : TQuery;
begin
  FContexte.LConvCpte := TList.Create;
  FContexte.LConvJal := TList.Create;
  FContexte.LCollectif := TList.Create;
  FContexte.DebEx := VH^.Encours.Deb;
  FContexte.FinEx := VH^.Encours.Fin;
  FContexte.EncEx := VH^.Encours.Fin;
  FContexte.Base := THValComboBox(GetControl('ZBase')).Value;
  FContexte.FileName := GetControlText('ZFichier');
end;

procedure TOF_RECUPDONNEES.LibereContexte;
begin
  FContexte.LConvCpte.Free;
  FContexte.LConvJal.Free;
  FContexte.LCollectif.Free;
end;

Initialization
  registerclasses ( [ TOF_RECUPDONNEES ] ) ;
end.
