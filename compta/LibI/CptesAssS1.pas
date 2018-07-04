unit CptesAssS1;

// 07/07/1999 - CA - Gestion du dernier paramètre dans lookup

interface

uses
  Windows
  {$IFDEF VER150}
  , variants
  {$ENDIF VER150}
{$IFNDEF EAGLCLIENT}
  ,db
    {$IFDEF ODBCDAC}
    , odbcconnection, odbctable, odbcquery, odbcdac
    {$ELSE}
      {$IFNDEF DBXPRESS}
      ,dbtables
      {$ELSE}
      ,uDbxDataSet
      {$ENDIF}
    {$ENDIF}
  {$ELSE}
  {$ENDIF EAGLCLIENT}
  , utob ,
  Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, StdCtrls, Hctrls, HSysMenu, hmsgbox,
   Hqry, HTB97, DBCtrls, ExtCtrls, Grids, DBGrids, HDB, HEnt1,
  Mask, HPanel, UiUtil,LookUp, ImEnt, imoutgen,ContCpte, ADODB ;

  function VerifCompteAssocie (Sender : TObject) : boolean;
  Procedure ConsultationComptesAsso(Comment : TActionFiche);
  function TesteExistenceUnCompte (Compte : string) : boolean;
  procedure AjouteCompteDansListe ( var L : TList; Compte : string);
  procedure OnEllipsisCompteImmoGene (Sender : TObject);

type

  TFComptesAssocies = class(TFFicheListe)
    TaIC_COMPTEIMMO: TStringField;
    TaIC_COMPTEAMORT: TStringField;
    TaIC_COMPTEDOTATION: TStringField;
    TaIC_COMPTEDEROG: TStringField;
    TaIC_REPRISEDEROG: TStringField;
    TaIC_DOTATIONEXC: TStringField;
    TaIC_VACEDEE: TStringField;
    TaIC_AMORTCEDE: TStringField;
    TaIC_VOACEDE: TStringField;
    TaIC_PROVISDEROG: TStringField;
    HPanel1: THPanel;
    HLabel1: THLabel;
    HPanel2: THPanel;
    HLabel2: THLabel;
    HLabel4: THLabel;
    HLabel11: THLabel;
    TaIC_REPEXPLOIT: TStringField;
    HPanel3: THPanel;
    HLabel8: THLabel;
    HLabel12: THLabel;
    HPanel4: THPanel;
    HLabel9: THLabel;
    HLabel10: THLabel;
    HLabel3: THLabel;
    HPanel5: THPanel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    HLabel5: THLabel;
    TaIC_REPEXCEP: TStringField;
    IC_COMPTEAMORT: THDBEdit;
    IC_COMPTEDOTATION: THDBEdit;
    IC_REPEXPLOIT: THDBEdit;
    IC_DOTATIONEXC: THDBEdit;
    IC_REPEXCEP: THDBEdit;
    IC_VACEDEE: THDBEdit;
    IC_AMORTCEDE: THDBEdit;
    IC_VOACEDE: THDBEdit;
    IC_COMPTEDEROG: THDBEdit;
    IC_PROVISDEROG: THDBEdit;
    IC_REPRISEDEROG: THDBEdit;
    IC_COMPTEIMMO: THDBEdit;
    TGene: TTable;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure OnCompteExit(Sender: TObject);
    Function  EnregOK : boolean ; Override ;
    procedure OnEllipsisCompteImmo(Sender: TObject);
    procedure OnCompteEnter(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    TypeAction : TActionFiche;
    fCurCompte : string;
  public
    { Déclarations publiques }
  end;

function TesteExistenceComptesAssocies (Ass : TCompteAss; var L : TList) : boolean;
function VerifEtCreationComptesAssocies (CpteAssocies : TCompteAss) : boolean;

implementation

{$R *.DFM}
{$IFDEF SERIE1} //YCP 25/08/05
uses S1Util ;
{$ENDIF SERIE1}

Procedure ConsultationComptesAsso(Comment : TActionFiche);
var FComptesAssocies: TFComptesAssocies; PP: THPanel ;
begin
   FComptesAssocies:=TFComptesAssocies.Create(Application) ;
   FComptesAssocies.TypeAction:=Comment ;
   FComptesAssocies.TGene.Open;
   FComptesAssocies.InitFL('IC','PRT_IMMOCPTE','','',Comment,TRUE,FComptesAssocies.TaIC_COMPTEIMMO,
   FComptesAssocies.TaIC_COMPTEIMMO,FComptesAssocies.TaIC_COMPTEIMMO,['']) ;
   {$IFDEF SERIE1}
   FComptesAssocies.HelpContext:=521000 ;
   {$ELSE}
   {$ENDIF}

   PP:=FindInsidePanel ;
   if PP=Nil then
   begin
     try     FComptesAssocies.ShowModal ;
     finally FComptesAssocies.Free ;
     end ;
   end
   else
   begin
     InitInside(FComptesAssocies,PP) ;
     FComptesAssocies.Show ;
   end;
   Screen.Cursor:=SyncrDefault ;
end;

function TFComptesAssocies.EnregOK : boolean ;
var ComptesAssocies : TCompteAss;
    Q : TQuery;
    nbEnreg : longint;
    MessageInfo : string;
    i,j : integer;
begin
  ComptesAssocies.Immo           :=TaIC_COMPTEIMMO.AsString;
  ComptesAssocies.Amort          :=TaIC_COMPTEAMORT.AsString;
  ComptesAssocies.Dotation       :=TaIC_COMPTEDOTATION.AsString;
  ComptesAssocies.Derog          :=TaIC_COMPTEDEROG.AsString;
  ComptesAssocies.RepriseDerog   :=TaIC_REPRISEDEROG.AsString;
  ComptesAssocies.DotationExcep  :=TaIC_DOTATIONEXC.AsString;
  ComptesAssocies.VaCedee        :=TaIC_VACEDEE.AsString;
  ComptesAssocies.AmortCede      :=TaIC_AMORTCEDE.AsString;
  ComptesAssocies.VoaCede        :=TaIC_VOACEDE.AsString;
  ComptesAssocies.ProvisDerog    :=TaIC_PROVISDEROG.AsString;
  ComptesAssocies.RepriseExcep   :=TaIC_REPEXCEP.AsString;
  ComptesAssocies.RepriseExploit :=TaIC_REPEXPLOIT.AsString;
  for i:=0 to PAppli.ControlCount - 1 do
  begin
    if (PAppli.Controls[i] is THPanel) then
    begin
      for j:=0 to THPanel(PAppli.Controls[i]).ControlCount - 1 do
      begin
        if (THPanel(PAppli.Controls[i]).Controls[j] is THDBEdit) then
        begin
          if (not VerifCompteAssocie (THPanel(PAppli.Controls[i]).Controls[j])) then
          begin
            HM.Execute(13,Caption,'');
            result := false;
            exit;
          end;
        end;
      end;
    end;
  end;
  Result := VerifEtCreationComptesAssocies (ComptesAssocies);
  if not Result then Exit;
  Result := inherited EnregOK;
  if Result then
  begin
    Q:= OpenSQL ('SELECT I_IMMO FROM IMMO WHERE I_NATUREIMMO="PRO" AND I_COMPTEIMMO="'+ComptesAssocies.Immo+'"',True);
    if not Q.Eof then
    begin
      if HM.Execute (8,HM.Mess[12],'')= mrYes then
      begin
        nbEnreg := ExecuteSQL('UPDATE IMMO SET I_COMPTEAMORT="'+ComptesAssocies.Amort+
                           '",I_COMPTEDOTATION="'+ComptesAssocies.Dotation+
                           '",I_COMPTEDEROG="'+ComptesAssocies.Derog+
                           '",I_REPRISEDEROG="'+ComptesAssocies.RepriseDerog+
                           '",I_PROVISDEROG="'+ComptesAssocies.ProvisDerog+
                           '",I_DOTATIONEXC="'+ComptesAssocies.DotationExcep+
                           '",I_VACEDEE="'+ComptesAssocies.VaCedee+
                           '",I_AMORTCEDE="'+ComptesAssocies.AmortCede+
                           '",I_REPEXPLOIT="'+ComptesAssocies.RepriseExploit+
                           '",I_REPEXCEP="'+ComptesAssocies.RepriseExcep+
                           '",I_VAOACEDEE="'+ComptesAssocies.VoaCede+
                           '" WHERE I_COMPTEIMMO="'+ComptesAssocies.Immo+'"');
        if nbEnreg <= 1 then MessageInfo := IntToStr(nbEnreg)+' '+HM.Mess[9]
        else MessageInfo := IntToStr(nbEnreg)+' '+HM.Mess[10];
        HM.Execute(11,HM.Mess[12],MessageInfo);
      end;
    end;
    Ferme(Q);
  end;
end;

procedure TFComptesAssocies.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  TGene.Close;
  inherited;
  if isInside(Self) then Action:=caFree ;
end;



procedure TFComptesAssocies.OnCompteExit(Sender: TObject);
var Compte,CompteImmo:string;
begin
  inherited;
  if (Ta.State = dsBrowse) then exit;
  if THDBEdit(Sender).Field.Value = NULL then exit;
  Compte := THDBEdit(Sender).Field.Value;
  if Compte = fCurCompte then exit;
  if Compte = '' then exit;
  if Presence('IMMOCPTE','IC_COMPTEIMMO',Compte) then exit;
  Compte := ImBourreEtLess ( Compte,fbGene);
  if Presence('GENERAUX','G_GENERAL',Compte) then
  begin
    THDBEdit(Sender).Field.Value := Compte;
  end;
  //  Proposition de comptes
  if THDBEdit(Sender).Name = 'IC_COMPTEIMMO' then
  begin
    if TaIC_COMPTEIMMO.AsString='' then exit; //EPZ 28/10/98
    CompteImmo := TaIC_COMPTEIMMO.AsString;
    if CompteImmo[1] = '2' then
    begin
      TaIC_COMPTEAMORT.AsString := CAssAmortissement (CompteImmo);
      if not VerifCompteAssocie(IC_COMPTEAMORT) then TaIC_COMPTEAMORT.AsString:= '';
      TaIC_COMPTEDOTATION.AsString := CAssDotation (CompteImmo);
      if not VerifCompteAssocie(IC_COMPTEDOTATION) then TaIC_COMPTEDOTATION.AsString:= '';
      TaIC_COMPTEDEROG.AsString := CAssDerog(CompteImmo);
      if not VerifCompteAssocie(IC_COMPTEDEROG) then TaIC_COMPTEDEROG.AsString:= '';
      TaIC_REPRISEDEROG.AsString := CAssRepriseDerog(CompteImmo);
      if not VerifCompteAssocie(IC_REPRISEDEROG) then TaIC_REPRISEDEROG.AsString:= '';
      TaIC_PROVISDEROG.AsString := CAssProvisDerog(CompteImmo);
      if not VerifCompteAssocie(IC_PROVISDEROG) then TaIC_PROVISDEROG.AsString:= '';
      TaIC_DOTATIONEXC.AsString := CAssDotationExc (TaIC_COMPTEDOTATION.AsString);
      if not VerifCompteAssocie(IC_DOTATIONEXC) then TaIC_DOTATIONEXC.AsString:= '';
      TaIC_VACEDEE.AsString := CAssVaCedee(CompteImmo);
      if not VerifCompteAssocie(IC_VACEDEE) then TaIC_VACEDEE.AsString:= '';
      TaIC_AMORTCEDE.AsString := TaIC_COMPTEAMORT.AsString;
      if not VerifCompteAssocie(IC_AMORTCEDE) then TaIC_AMORTCEDE.AsString:= '';
      TaIC_VOACEDE.AsString := TaIC_COMPTEIMMO.AsString;
      if not VerifCompteAssocie(IC_VOACEDE) then TaIC_VOACEDE.AsString:= '';
      TaIC_REPEXPLOIT.AsString := CAssRepExploit (TaIC_COMPTEDOTATION.AsString);
      if not VerifCompteAssocie(IC_REPEXPLOIT) then TaIC_REPEXPLOIT.AsString:= '';
      TaIC_REPEXCEP.AsString := CAssRepExc (TaIC_DOTATIONEXC.AsString);
      if not VerifCompteAssocie(IC_REPEXCEP) then TaIC_REPEXCEP.AsString:= '';
    end;
  end;
end;

function TesteExistenceUnCompte (Compte : string) : boolean;
begin
  if Compte = '' then begin Result := true;exit;end;
  Result :=  Presence ( 'GENERAUX','G_GENERAL',Compte );
end;

procedure AjouteCompteDansListe ( var L : TList; Compte : string);
var ARecord : ^TDefCompte;
    i : integer;
    bTrouve : boolean;
begin
  if Compte = '' then exit;
  bTrouve := false;
  if L = nil then L := TList.Create;
  for i:=0 to L.Count-1 do
  begin
    ARecord := L.Items[i];
    if ARecord^.Compte = Compte then
    begin
      bTrouve := True;
      break;
    end;
  end;
  if bTrouve = False then
  begin
    new (ARecord);
    ARecord^.Compte := Compte;
    ARecord^.Libelle := '';
    L.Add(ARecord);
  end;
end;

function TesteExistenceComptesAssocies (Ass : TCompteAss; var L : TList) : boolean;
begin
  if not TesteExistenceUnCompte(Ass.Immo)  then
    AjouteCompteDansListe (L,Ass.Immo);
  if not TesteExistenceUnCompte(Ass.Amort) then
    AjouteCompteDansListe (L,Ass.Amort);
  if not TesteExistenceUnCompte(Ass.Dotation) then
    AjouteCompteDansListe (L,Ass.Dotation);
  if not TesteExistenceUnCompte(Ass.Derog) then
    AjouteCompteDansListe (L,Ass.Derog);
  if not TesteExistenceUnCompte(Ass.RepriseDerog) then
    AjouteCompteDansListe (L,Ass.RepriseDerog);
  if not TesteExistenceUnCompte(Ass.DotationExcep) then
    AjouteCompteDansListe (L,Ass.DotationExcep);
  if not TesteExistenceUnCompte(Ass.VaCedee) then
    AjouteCompteDansListe (L,Ass.VaCedee);
  if not TesteExistenceUnCompte(Ass.AmortCede) then
    AjouteCompteDansListe (L,Ass.AmortCede);
  if not TesteExistenceUnCompte(Ass.VoaCede) then
    AjouteCompteDansListe (L,Ass.VoaCede);
  if not TesteExistenceUnCompte(Ass.ProvisDerog) then
    AjouteCompteDansListe (L,Ass.ProvisDerog);
  if not TesteExistenceUnCompte(Ass.RepriseExcep) then
    AjouteCompteDansListe (L,Ass.RepriseExcep);
  if not TesteExistenceUnCompte(Ass.RepriseExploit) then
    AjouteCompteDansListe (L,Ass.RepriseExploit);
  if L = nil then Result := True else Result := False;
end;

// Cette fonction vérifie que les comptes associés ont bien été créés et
// propose leur création dans le cas contraire.
// Si après proposition de création, les comptes sont tous créés : Result := True
// Sinon Result := False;

function VerifEtCreationComptesAssocies (CpteAssocies : TCompteAss) : boolean;
var ListeCompte : TList;  i: integer ;ARecord : ^TDefCompte;
begin
  ListeCompte := nil;
  if not TesteExistenceComptesAssocies (CpteAssocies,ListeCompte) then
  begin
    CreationComptesImmo (ListeCompte);
    if ListeCompte = nil then Result := True
    else
    begin
      for i:=0 to ListeCompte.Count - 1 do
      begin
        ARecord := ListeCompte.Items[i];
        Dispose(ARecord);
      end;
      ListeCompte.Free;
      ListeCompte := nil;
      Result := false;
    end;
  end
  else
    Result := True;
end;

procedure TFComptesAssocies.OnEllipsisCompteImmo(Sender: TObject);
begin
  inherited;
  OnEllipsisCompteImmoGene (Sender);
end;

procedure TFComptesAssocies.OnCompteEnter(Sender: TObject);
begin
  inherited;
  fCurCompte :=THDBEdit(Sender).Field.Value;
end;

procedure OnEllipsisCompteImmoGene (Sender : TObject);
var stWhere,TypeCompte : string;
begin
  TypeCompte := ExtractSuffixe(THDBEdit(Sender).Name);
  if TypeCompte = 'COMPTEIMMO' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteImmoSup +'" AND G_GENERAL>="'+VHImmo^.CpteImmoInf+'"'
  else   if TypeCompte = 'COMPTEAMORT' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteAmortSup +'" AND G_GENERAL>="'+VHImmo^.CpteAmortInf+'"'
  else   if TypeCompte = 'COMPTEDOTATION' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteDotSup +'" AND G_GENERAL>="'+VHImmo^.CpteDotInf+'"'
  else   if TypeCompte = 'REPEXPLOIT' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteExploitSup +'" AND G_GENERAL>="'+VHImmo^.CpteExploitInf +'"'
  else   if TypeCompte = 'DOTATIONEXC' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteDotExcSup +'" AND G_GENERAL>="'+VHImmo^.CpteDotExcInf +'"'
  else   if TypeCompte = 'REPEXCEP' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteRepExcSup +'" AND G_GENERAL>="'+VHImmo^.CpteRepExcInf +'"'
  else   if TypeCompte = 'VACEDEE' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteVaCedeeSup +'" AND G_GENERAL>="'+VHImmo^.CpteVaCedeeInf +'"'
  else   if TypeCompte = 'AMORTCEDE' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteAmortSup +'" AND G_GENERAL>="'+VHImmo^.CpteAmortInf +'"'
  else   if (TypeCompte = 'VOACEDE') or (TypeCompte = 'VAOACEDEE') then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteImmoSup +'" AND G_GENERAL>="'+VHImmo^.CpteImmoInf +'"'
  else   if TypeCompte = 'COMPTEDEROG' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteDerogSup +'" AND G_GENERAL>="'+VHImmo^.CpteDerogInf +'"'
  else   if TypeCompte = 'PROVISDEROG' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteProvDerSup +'" AND G_GENERAL>="'+VHImmo^.CpteProvDerInf +'"'
  else   if TypeCompte = 'REPRISEDEROG' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteRepDerSup +'" AND G_GENERAL>="'+VHImmo^.CpteRepDerInf +'"';
  {$ifdef S5}  // YCP 01/12/00
  LookupList(TControl(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,1)  ;
  {$else}
  LookupList(TControl(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,3)  ;
  {$endif}
end;

function VerifCompteAssocie (Sender : TObject) : boolean;
var stSup,stInf,TypeCompte : string;
begin
  TypeCompte := ExtractSuffixe(THDBEdit(Sender).Name);
  if TypeCompte = 'COMPTEIMMO' then
   begin stSup := VHImmo^.CpteImmoSup;stInf:=VHImmo^.CpteImmoInf; end
  else   if TypeCompte = 'COMPTEAMORT' then
     begin stSup := VHImmo^.CpteAmortSup;stInf:=VHImmo^.CpteAmortInf; end
  else   if TypeCompte = 'COMPTEDOTATION' then
     begin stSup := VHImmo^.CpteDotSup;stInf:=VHImmo^.CpteDotInf; end
  else   if TypeCompte = 'REPEXPLOIT' then
     begin stSup := VHImmo^.CpteExploitSup;stInf:=VHImmo^.CpteExploitInf; end
  else   if TypeCompte = 'DOTATIONEXC' then
     begin stSup := VHImmo^.CpteDotExcSup;stInf:=VHImmo^.CpteDotExcInf; end
  else   if TypeCompte = 'REPEXCEP' then
     begin stSup := VHImmo^.CpteRepExcSup;stInf:=VHImmo^.CpteRepExcInf; end
  else   if TypeCompte = 'VACEDEE' then
     begin stSup := VHImmo^.CpteVaCedeeSup;stInf:=VHImmo^.CpteVaCedeeInf; end
  else   if TypeCompte = 'AMORTCEDE' then
     begin stSup := VHImmo^.CpteAmortSup;stInf:=VHImmo^.CpteAmortInf; end
  else   if (TypeCompte = 'VOACEDE') or (TypeCompte = 'VAOACEDEE') OR (TypeCompte='VOACEDEE') then
     begin stSup := VHImmo^.CpteImmoSup;stInf:=VHImmo^.CpteImmoInf; end
  else   if TypeCompte = 'COMPTEDEROG' then
     begin stSup := VHImmo^.CpteDerogSup;stInf:=VHImmo^.CpteDerogInf; end
  else   if TypeCompte = 'PROVISDEROG' then
     begin stSup := VHImmo^.CpteProvDerSup;stInf:=VHImmo^.CpteProvDerInf; end
  else   if TypeCompte = 'REPRISEDEROG' then
     begin stSup := VHImmo^.CpteRepDerSup;stInf:=VHImmo^.CpteRepDerInf; end;
  result := ((THDBEdit(Sender).Field.AsString <= stSup) and (THDBEdit(Sender).Field.AsString >= stInf))
end;

procedure TFComptesAssocies.FormShow(Sender: TObject);
begin
  inherited;
  if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
  BEGIN
    if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
  END ;
//EPZ 30/10/00
  IC_COMPTEAMORT.MaxLength    := VHImmo^.Cpta[fbGene].Lg;
  IC_COMPTEDOTATION.MaxLength := VHImmo^.Cpta[fbGene].Lg;
  IC_REPEXPLOIT.MaxLength     := VHImmo^.Cpta[fbGene].Lg;
  IC_DOTATIONEXC.MaxLength    := VHImmo^.Cpta[fbGene].Lg;
  IC_REPEXCEP.MaxLength       := VHImmo^.Cpta[fbGene].Lg;
  IC_VACEDEE.MaxLength        := VHImmo^.Cpta[fbGene].Lg;
  IC_AMORTCEDE.MaxLength      := VHImmo^.Cpta[fbGene].Lg;
  IC_VOACEDE.MaxLength        := VHImmo^.Cpta[fbGene].Lg;
  IC_COMPTEDEROG.MaxLength    := VHImmo^.Cpta[fbGene].Lg;
  IC_PROVISDEROG.MaxLength    := VHImmo^.Cpta[fbGene].Lg;
  IC_REPRISEDEROG.MaxLength   := VHImmo^.Cpta[fbGene].Lg;
  IC_COMPTEIMMO.MaxLength     := VHImmo^.Cpta[fbGene].Lg;
//EPZ 30/10/00
end;


(*////////////////////////////////////////////////////////////////////////////////////////////////////////


           FONCTIONS A EXTERNALISER

////////////////////////////////////////////////////////////////////////////////////////////////////////*)


end.


