unit CptesAss;

// 07/07/1999 - CA - Gestion du dernier paramètre dans lookup

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, StdCtrls, Hctrls, Db, HSysMenu, hmsgbox,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$IFDEF VER150}
   Variants,
  {$ENDIF}
  Hqry, HTB97, DBCtrls, ExtCtrls, Grids, DBGrids, HDB, HEnt1,
  Mask, HPanel,UiUtil,LookUp,ImEnt, ImOutGen ;

  Procedure ConsultationComptesAsso(Comment : TActionFiche);
  function TesteExistenceUnCompte (Compte : string) : boolean;
  procedure AjouteCompteDansListe ( var L : TList; Compte : string);
  procedure OnEllipsisCompteImmoGene (Sender : TObject);
type
  TFComptesAssocies = class(TFFicheListe)
    TaPC_COMPTEIMMO: TStringField;
    TaPC_COMPTEAMORT: TStringField;
    TaPC_COMPTEDOTATION: TStringField;
    TaPC_COMPTEDEROG: TStringField;
    TaPC_REPRISEDEROG: TStringField;
    TaPC_DOTATIONEXC: TStringField;
    TaPC_VACEDEE: TStringField;
    TaPC_AMORTCEDE: TStringField;
    TaPC_VOACEDE: TStringField;
    TaPC_PROVISDEROG: TStringField;
    HPanel1: THPanel;
    HLabel1: THLabel;
    HPanel2: THPanel;
    HLabel2: THLabel;
    HLabel4: THLabel;
    HLabel11: THLabel;
    TaPC_REPEXPLOIT: TStringField;
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
    TaPC_REPEXCEP: TStringField;
    PC_COMPTEAMORT: THDBEdit;
    PC_COMPTEDOTATION: THDBEdit;
    PC_REPEXPLOIT: THDBEdit;
    PC_DOTATIONEXC: THDBEdit;
    PC_REPEXCEP: THDBEdit;
    PC_VACEDEE: THDBEdit;
    PC_AMORTCEDE: THDBEdit;
    PC_VOACEDE: THDBEdit;
    PC_COMPTEDEROG: THDBEdit;
    PC_PROVISDEROG: THDBEdit;
    PC_REPRISEDEROG: THDBEdit;
    PC_COMPTEIMMO: THDBEdit;
    TGene: TTable;
    TaPC_LIBELLE: TStringField;
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

procedure RecupereComptesAssocies(QImmo : TQuery; RefCompteImmo : string;var ARecord : TCompteAss);
function TesteExistenceComptesAssocies (Ass : TCompteAss; var L : TList) : boolean;
function VerifEtCreationComptesAssocies (CpteAssocies : TCompteAss) : boolean;

implementation

uses ContCpte,integecr;

{$R *.DFM}

Procedure ConsultationComptesAsso(Comment : TActionFiche);
var FComptesAssocies: TFComptesAssocies;
    PP : THPanel ;
begin
FComptesAssocies:=TFComptesAssocies.Create(Application) ;
FComptesAssocies.TypeAction:=Comment ;
FComptesAssocies.TGene.Open;
FComptesAssocies.InitFL('PC','PRT_IMMOCPTE','','',Comment,TRUE,FComptesAssocies.TaPC_COMPTEIMMO,
FComptesAssocies.TaPC_LIBELLE,FComptesAssocies.TaPC_COMPTEIMMO,['']) ;
PP:=FindInsidePanel ;
if PP=Nil then
  begin
  try FComptesAssocies.ShowModal ; finally FComptesAssocies.Free ; end ;
  end else
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
  ComptesAssocies.Immo :=TaPC_COMPTEIMMO.AsString;
  ComptesAssocies.Amort :=TaPC_COMPTEAMORT.AsString;
  ComptesAssocies.Dotation :=TaPC_COMPTEDOTATION.AsString;
  ComptesAssocies.Derog :=TaPC_COMPTEDEROG.AsString;
  ComptesAssocies.RepriseDerog :=TaPC_REPRISEDEROG.AsString;
  ComptesAssocies.DotationExcep :=TaPC_DOTATIONEXC.AsString;
  ComptesAssocies.VaCedee :=TaPC_VACEDEE.AsString;
  ComptesAssocies.AmortCede :=TaPC_AMORTCEDE.AsString;
  ComptesAssocies.VoaCede :=TaPC_VOACEDE.AsString;
  ComptesAssocies.ProvisDerog :=TaPC_PROVISDEROG.AsString;
  ComptesAssocies.RepriseExcep :=TaPC_REPEXCEP.AsString;
  ComptesAssocies.RepriseExploit :=TaPC_REPEXPLOIT.AsString;

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

procedure RecupereComptesAssocies(QImmo : TQuery; RefCompteImmo : string;var ARecord : TCompteAss);
var QComptesAsso : TQuery; CompteImmo,Prefix : string;
begin
  {$IFDEF SERIE1}
  Prefix:='IC' ;
  {$ELSE}
  Prefix:='PC' ;
  {$ENDIF}
  CompteImmo:=RefCompteImmo;
  if (QImmo <> nil) then
  begin
    ARecord.Immo := QImmo.FindField('I_COMPTEIMMO').AsString;
    ARecord.Amort := QImmo.FindField('I_COMPTEAMORT').AsString;
    ARecord.Dotation := QImmo.FindField('I_COMPTEDOTATION').AsString;
    ARecord.Derog := QImmo.FindField('I_COMPTEDEROG').AsString;
    ARecord.RepriseDerog := QImmo.FindField('I_REPRISEDEROG').AsString;
    ARecord.DotationExcep := QImmo.FindField('I_DOTATIONEXC').AsString;
    ARecord.ProvisDerog := QImmo.FindField('I_PROVISDEROG').AsString;
    ARecord.VaCedee := QImmo.FindField('I_VACEDEE').AsString;
    ARecord.AmortCede := QImmo.FindField('I_AMORTCEDE').AsString;
    ARecord.VoaCede := QImmo.FindField('I_VAOACEDEE').AsString;
    ARecord.RepriseExcep := QImmo.FindField('I_REPEXCEP').AsString;
    ARecord.RepriseExploit := QImmo.FindField('I_REPEXPLOIT').AsString;
//EPZ 28/11/00
{ A FAIRE
    ARecord.CreanceCessionActif := QImmo.FindField('I_CREANCESURCESS').AsString;
    ARecord.ProduitCessionActif := QImmo.FindField('I_PDTCESSACTIF').AsString;
    ARecord.TVACollectee := QImmo.FindField('I_TVACOLLECTEE').AsString;
}
//EPZ 28/11/00
  end
  else
  begin
//EPZ 30/10/00
    if (CompteImmo = '') and (QImmo <> nil) then
      CompteImmo := QImmo.FindField('I_COMPTEIMMO').AsString;
    ImBourreLaDoncSurLesComptes(CompteImmo);
//EPZ 30/10/00
    QComptesAsso := OpenSQL('SELECT * FROM IMMOCPTE WHERE '+Prefix+'_COMPTEIMMO = "'+ CompteImmo +'"',TRUE);
    if Not QComptesAsso.EOF then
    begin
      QComptesAsso.First;
      ARecord.Immo := CompteImmo;
      ARecord.Amort := QComptesAsso.FindField(''+Prefix+'_COMPTEAMORT').AsString;
      ARecord.Dotation := QComptesAsso.FindField(''+Prefix+'_COMPTEDOTATION').AsString;
      ARecord.Derog := QComptesAsso.FindField(''+Prefix+'_COMPTEDEROG').AsString;
      ARecord.RepriseDerog := QComptesAsso.FindField(''+Prefix+'_REPRISEDEROG').AsString;
      ARecord.DotationExcep := QComptesAsso.FindField(''+Prefix+'_DOTATIONEXC').AsString;
      ARecord.ProvisDerog := QComptesAsso.FindField(''+Prefix+'_PROVISDEROG').AsString;
      ARecord.VaCedee := QComptesAsso.FindField(''+Prefix+'_VACEDEE').AsString;
      ARecord.AmortCede := QComptesAsso.FindField(''+Prefix+'_AMORTCEDE').AsString;
      ARecord.VoaCede := QComptesAsso.FindField(''+Prefix+'_VOACEDE').AsString;
      ARecord.RepriseExcep := QComptesAsso.FindField(''+Prefix+'_REPEXCEP').AsString;
      ARecord.RepriseExploit := QComptesAsso.FindField(''+Prefix+'_REPEXPLOIT').AsString;
//EPZ 28/11/00
{ A FAIRE
      ARecord.CreanceCessionActif := QComptesAsso.FindField('PC_CREANCESURCESS').AsString;
      ARecord.ProduitCessionActif := QComptesAsso.FindField('PC_PDTCESSACTIF').AsString;
      ARecord.TVACollectee := QComptesAsso.FindField('PC_TVACOLLECTEE').AsString;
}
//EPZ 28/11/00
    end
    else
    begin
      GetComptesAssociesParDefaut(ARecord,CompteImmo) ;
    end;
    Ferme (QComptesAsso);
  end;
//EPZ 09/11/00  result := ARecord;
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
  if Presence('IMMOCPTE','PC_COMPTEIMMO',Compte) then exit;
  Compte := ImBourreEtLess ( Compte,ImGeneTofb);
  if Presence('GENERAUX','G_GENERAL',Compte) then
  begin
    THDBEdit(Sender).Field.Value := Compte;
  end;
  //  Proposition de comptes
  if THDBEdit(Sender).Name = 'PC_COMPTEIMMO' then
  begin
    if TaPC_COMPTEIMMO.AsString='' then exit; //EPZ 28/10/98
    CompteImmo := TaPC_COMPTEIMMO.AsString;
    if CompteImmo[1] = '2' then
    begin
      TaPC_COMPTEAMORT.AsString := CAssAmortissement (CompteImmo);
      if not VerifCompteAssocie(PC_COMPTEAMORT) then TaPC_COMPTEAMORT.AsString:= '';
      TaPC_COMPTEDOTATION.AsString := CAssDotation (CompteImmo);
      if not VerifCompteAssocie(PC_COMPTEDOTATION) then TaPC_COMPTEDOTATION.AsString:= '';
      TaPC_COMPTEDEROG.AsString := CAssDerog(CompteImmo);
      if not VerifCompteAssocie(PC_COMPTEDEROG) then TaPC_COMPTEDEROG.AsString:= '';
      TaPC_REPRISEDEROG.AsString := CAssRepriseDerog(CompteImmo);
      if not VerifCompteAssocie(PC_REPRISEDEROG) then TaPC_REPRISEDEROG.AsString:= '';
      TaPC_PROVISDEROG.AsString := CAssProvisDerog(CompteImmo);
      if not VerifCompteAssocie(PC_PROVISDEROG) then TaPC_PROVISDEROG.AsString:= '';
      TaPC_DOTATIONEXC.AsString := CAssDotationExc (TaPC_COMPTEDOTATION.AsString);
      if not VerifCompteAssocie(PC_DOTATIONEXC) then TaPC_DOTATIONEXC.AsString:= '';
      TaPC_VACEDEE.AsString := CAssVaCedee(CompteImmo);
      if not VerifCompteAssocie(PC_VACEDEE) then TaPC_VACEDEE.AsString:= '';
      TaPC_AMORTCEDE.AsString := TaPC_COMPTEAMORT.AsString;
      if not VerifCompteAssocie(PC_AMORTCEDE) then TaPC_AMORTCEDE.AsString:= '';
      TaPC_VOACEDE.AsString := TaPC_COMPTEIMMO.AsString;
      if not VerifCompteAssocie(PC_VOACEDE) then TaPC_VOACEDE.AsString:= '';
      TaPC_REPEXPLOIT.AsString := CAssRepExploit (TaPC_COMPTEDOTATION.AsString);
      if not VerifCompteAssocie(PC_REPEXPLOIT) then TaPC_REPEXPLOIT.AsString:= '';
      TaPC_REPEXCEP.AsString := CAssRepExc (TaPC_DOTATIONEXC.AsString);
      if not VerifCompteAssocie(PC_REPEXCEP) then TaPC_REPEXCEP.AsString:= '';
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
  if not TesteExistenceUnCompte(Ass.Immo) then
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
//EPZ 28/11/00
{ A FAIRE
  if not TesteExistenceUnCompte(Ass.CreanceCessionActif) then
    AjouteCompteDansListe (L,Ass.CreanceCessionActif);
  if not TesteExistenceUnCompte(Ass.ProduitCessionActif) then
    AjouteCompteDansListe (L,Ass.ProduitCessionActif);
  if not TesteExistenceUnCompte(Ass.TVACollectee) then
    AjouteCompteDansListe (L,Ass.TVACollectee);
}
//EPZ 28/11/00

  if L = nil then Result := True else Result := False;
end;

// Cette fonction vérifie que les comptes associés ont bien été créés et
// propose leur création dans le cas contraire.
// Si après proposition de création, les comptes sont tous créés : Result := True
// Sinon Result := False;

function VerifEtCreationComptesAssocies (CpteAssocies : TCompteAss) : boolean;
var i : integer;
    ListeCompte : TList;
    ARecord : ^TDefCompte;
begin
  ListeCompte := nil;
  if not TesteExistenceComptesAssocies(CpteAssocies,ListeCompte) then
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
  end else Result := True;
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
//EPZ 28/11/00
{ A FAIRE
  else   if TypeCompte = 'CREANCESURCESS' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteCreanCessActSup +'" AND G_GENERAL>="'+VHImmo^.CpteCreanCessActInf +'"'
  else   if TypeCompte = 'PDTCESSACTIF' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CptePdtCessActSup +'" AND G_GENERAL>="'+VHImmo^.CptePdtCessActInf +'"'
  else   if TypeCompte = 'TVACOLLECTEE' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteTVACollecteeSup +'" AND G_GENERAL>="'+VHImmo^.CpteTVACollecteeInf +'"'
}
//EPZ 28/11/00
  else   if TypeCompte = 'REPRISEDEROG' then
    stWhere := 'G_GENERAL<="'+VHImmo^.CpteRepDerSup +'" AND G_GENERAL>="'+VHImmo^.CpteRepDerInf +'"';

  LookupList(TControl(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,1)  ;
end;

procedure TFComptesAssocies.FormShow(Sender: TObject);
begin
  inherited;
  if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
  BEGIN
    if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
  END ;
//EPZ 30/10/00
  PC_COMPTEAMORT.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_COMPTEDOTATION.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_REPEXPLOIT.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_DOTATIONEXC.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_REPEXCEP.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_VACEDEE.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_AMORTCEDE.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_VOACEDE.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_COMPTEDEROG.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_PROVISDEROG.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_REPRISEDEROG.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_COMPTEIMMO.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
//EPZ 28/11/00
{ A FAIRE
  PC_CREANCESURCESS.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_PDTCESSACTIF.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
  PC_TVACOLLECTEE.MaxLength := VHImmo^.Cpta[ImGeneTofb].Lg;
}
//EPZ 28/11/00
//EPZ 30/10/00
end;

end.

