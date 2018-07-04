{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/02/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : IMMOCPTE (IMMOCPTE)
Mots clefs ... : TOM;IMMOCPTE
*****************************************************************}
Unit IMMOCPTE_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     hdb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FichList,
     fe_main,
{$else}
     eFichList,
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob,
     LookUp,
     ImEnt,
     ImOutGen,
     ContCpte ;

const
  {$IFDEF SERIE1}
    PREFIXE_IMMOCPTE = 'IC'
  {$ELSE}
    PREFIXE_IMMOCPTE = 'PC'
  {$ENDIF}
  ;

Type
  TOM_IMMOCPTE = Class (TOM)
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
      {$IFDEF EAGLCLIENT}
      PC_COMPTEIMMO : THEdit;
      PC_COMPTEAMORT : THEdit;
      PC_COMPTEDOTATION : THEdit;
      PC_COMPTEDEROG : THEdit;
      PC_REPRISEDEROG : THEdit;
      PC_DOTATIONEXC : THEdit;
      PC_PROVISDEROG : THEdit;
      PC_VACEDEE : THEdit;
      PC_AMORTCEDE : THEdit;
      PC_VOACEDE : THEdit;
      PC_REPEXCEP : THEdit;
      PC_REPEXPLOIT : THEdit;
      {$ELSE}
      PC_COMPTEIMMO : THDBEdit;
      PC_COMPTEAMORT : THDBEdit;
      PC_COMPTEDOTATION : THDBEdit;
      PC_COMPTEDEROG : THDBEdit;
      PC_REPRISEDEROG : THDBEdit;
      PC_DOTATIONEXC : THDBEdit;
      PC_PROVISDEROG : THDBEdit;
      PC_VACEDEE : THDBEdit;
      PC_AMORTCEDE : THDBEdit;
      PC_VOACEDE : THDBEdit;
      PC_REPEXCEP : THDBEdit;
      PC_REPEXPLOIT : THDBEdit;
      {$ENDIF}
      procedure OnElipsisCompteClick ( Sender : TObject );
    end ;

procedure AMLanceFiche_ComptesAssocies ( LeMode : TActionFiche );

procedure RecupereComptesAssocies(QImmo : TQuery; RefCompteImmo : string;var ARecord : TCompteAss);
procedure AjouteCompteDansListe ( var L : TList; Compte : string);
procedure OnEllipsisCompteImmoGene (Sender : TObject);
function TesteExistenceUnCompte (Compte : string) : boolean;
function TesteExistenceComptesAssocies (Ass : TCompteAss; var L : TList) : boolean;
function VerifEtCreationComptesAssocies (CpteAssocies : TCompteAss) : boolean;

Implementation

{$IFDEF CCSTD}
{$ELSE}
uses IMMO_TOM ;
{$ENDIF}

const
      MAX_CPTEASS = 12;
      A_CPTESASS : array [1..MAX_CPTEASS] of string =
        (PREFIXE_IMMOCPTE+'_COMPTEIMMO',PREFIXE_IMMOCPTE+'_COMPTEAMORT',PREFIXE_IMMOCPTE+'_COMPTEDOTATION',PREFIXE_IMMOCPTE+'_COMPTEDEROG'
        ,PREFIXE_IMMOCPTE+'_REPRISEDEROG',PREFIXE_IMMOCPTE+'_DOTATIONEXC',PREFIXE_IMMOCPTE+'_VACEDEE',PREFIXE_IMMOCPTE+'_AMORTCEDE',PREFIXE_IMMOCPTE+'_VOACEDE'
        ,PREFIXE_IMMOCPTE+'_PROVISDEROG',PREFIXE_IMMOCPTE+'_REPEXCEP',PREFIXE_IMMOCPTE+'_REPEXPLOIT');

procedure AMLanceFiche_ComptesAssocies ( LeMode : TActionFiche );
var
  stAction: string;
begin
  if LeMode = taCreat then
    stAction := 'ACTION=CREATION'
  else if LeMode = taModif then
    stAction := 'ACTION=MODIFICATION'
  else
    stAction := 'ACTION=CONSULTATION';
  AGLLanceFiche('AM', 'AMCPTEASS', '', '', stAction);
end;

procedure TOM_IMMOCPTE.OnNewRecord ;
begin
  Inherited ;
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
end ;

procedure TOM_IMMOCPTE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_IMMOCPTE.OnUpdateRecord ;
var ComptesAssocies : TCompteAss;
    i : integer;
begin
  Inherited ;
  ComptesAssocies.Immo            :=GetField(PREFIXE_IMMOCPTE+'_COMPTEIMMO');
  ComptesAssocies.Amort           :=GetField(PREFIXE_IMMOCPTE+'_COMPTEAMORT');
  ComptesAssocies.Dotation        :=GetField(PREFIXE_IMMOCPTE+'_COMPTEDOTATION');
  ComptesAssocies.Derog           :=GetField(PREFIXE_IMMOCPTE+'_COMPTEDEROG');
  ComptesAssocies.RepriseDerog    :=GetField(PREFIXE_IMMOCPTE+'_REPRISEDEROG');
  ComptesAssocies.DotationExcep   :=GetField(PREFIXE_IMMOCPTE+'_DOTATIONEXC');
  ComptesAssocies.VaCedee         :=GetField(PREFIXE_IMMOCPTE+'_VACEDEE');
  ComptesAssocies.AmortCede       :=GetField(PREFIXE_IMMOCPTE+'_AMORTCEDE');
  ComptesAssocies.VoaCede         :=GetField(PREFIXE_IMMOCPTE+'_VOACEDE');
  ComptesAssocies.ProvisDerog     :=GetField(PREFIXE_IMMOCPTE+'_PROVISDEROG');
  ComptesAssocies.RepriseExcep    :=GetField(PREFIXE_IMMOCPTE+'_REPEXCEP');
  ComptesAssocies.RepriseExploit  :=GetField(PREFIXE_IMMOCPTE+'_REPEXPLOIT');

  { Vérification des comptes saisis }
  for i:=1 to MAX_CPTEASS do
  begin
    {$IFDEF EAGLCLIENT}
    if (not VerifCompteAssocie (THEdit(GetControl(A_CPTESASS[i])))) then
    {$ELSE}
    if (not VerifCompteAssocie (THDBEdit(GetControl(A_CPTESASS[i])))) then
    {$ENDIF}
    begin
      LastError := -1;
      LastErrorMsg := 'Compte associé incorrect.';
      SetFocusControl (A_CPTESASS[i]);
      exit;
    end;
  end;
  if not VerifEtCreationComptesAssocies (ComptesAssocies) then
  begin
    LastError := -1;
    LastErrorMsg := 'Erreur en création des comptes associés.';
    SetFocusControl (A_CPTESASS[1]);
    exit;
  end;
end ;

procedure TOM_IMMOCPTE.OnAfterUpdateRecord ;
var Q: TQuery;
    nbEnreg : integer;
    stMessageInfo : string;
begin
  Inherited ;
  Q:= OpenSQL ('SELECT I_IMMO FROM IMMO WHERE I_NATUREIMMO="PRO" AND I_COMPTEIMMO="'+GetField(PREFIXE_IMMOCPTE+'_COMPTEIMMO')+'"',True);
  if not Q.Eof then
  begin
    if PGIAsk('Voulez-vous mettre à jour les fiches déjà créées ?') = mrYes then
    begin
      nbEnreg := ExecuteSQL('UPDATE IMMO SET I_COMPTEAMORT="'+GetField(PREFIXE_IMMOCPTE+'_COMPTEAMORT')+
                           '",I_COMPTEDOTATION="'+GetField(PREFIXE_IMMOCPTE+'_COMPTEDOTATION')+
                           '",I_COMPTEDEROG="'+GetField(PREFIXE_IMMOCPTE+'_COMPTEDEROG')+
                           '",I_REPRISEDEROG="'+GetField(PREFIXE_IMMOCPTE+'_REPRISEDEROG')+
                           '",I_PROVISDEROG="'+GetField(PREFIXE_IMMOCPTE+'_PROVISDEROG')+
                           '",I_DOTATIONEXC="'+GetField(PREFIXE_IMMOCPTE+'_DOTATIONEXC')+
                           '",I_VACEDEE="'+GetField(PREFIXE_IMMOCPTE+'_VACEDEE')+
                           '",I_AMORTCEDE="'+GetField(PREFIXE_IMMOCPTE+'_AMORTCEDE')+
                           '",I_REPEXPLOIT="'+GetField(PREFIXE_IMMOCPTE+'_REPEXPLOIT')+
                           '",I_REPEXCEP="'+GetField(PREFIXE_IMMOCPTE+'_REPEXCEP')+
                           '",I_VAOACEDEE="'+GetField(PREFIXE_IMMOCPTE+'_VOACEDE')+
                           '" WHERE I_COMPTEIMMO="'+GetField(PREFIXE_IMMOCPTE+'_COMPTEIMMO')+'"');
      if nbEnreg <= 1 then stMessageInfo := IntToStr(nbEnreg)+' '+TraduireMemoire('fiche a été mise à jour.')
      else stMessageInfo := IntToStr(nbEnreg)+' '+TraduireMemoire('fiches ont été mises à jour.');
      PGIInfo ( stMessageInfo );
    end;
    Ferme(Q);
  end;
end ;

procedure TOM_IMMOCPTE.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_IMMOCPTE.OnChangeField ( F: TField ) ;
var stCompte : string;
    CompteImmo : string;
    stNomChamp : string;
begin
  Inherited ;
  if DS.State in [dsBrowse] then exit;
  stNomChamp := F.FieldName;
  if GetField(PREFIXE_IMMOCPTE+'_COMPTEIMMO')='' then exit;
  stCompte := ImBourreEtLess ( GetField(stNomChamp),ImGeneTofb);
  { On bourre le compte avec les caractères de bourrage }
  if ((Presence('GENERAUX','G_GENERAL',stCompte)) and ( stCompte<>GetField(stNomChamp))) then
      SetField(stNomChamp, stCompte );

  { Comptes par défaut }
  if stNomChamp = PREFIXE_IMMOCPTE+'_COMPTEIMMO' then
  begin
    CompteImmo := GetField(PREFIXE_IMMOCPTE+'_COMPTEIMMO');
    if CompteImmo[1] = '2' then
    begin
      SetField(PREFIXE_IMMOCPTE+'_COMPTEAMORT',CAssAmortissement (CompteImmo));
      if not VerifCompteAssocie(PC_COMPTEAMORT) then SetField(PREFIXE_IMMOCPTE+'_COMPTEAMORT','');
      SetField(PREFIXE_IMMOCPTE+'_COMPTEDOTATION', CAssDotation (CompteImmo));
      if not VerifCompteAssocie(PC_COMPTEDOTATION) then SetField(PREFIXE_IMMOCPTE+'_COMPTEDOTATION','');
      SetField(PREFIXE_IMMOCPTE+'_COMPTEDEROG',CAssDerog(CompteImmo));
      if not VerifCompteAssocie(PC_COMPTEDEROG) then SetField(PREFIXE_IMMOCPTE+'_COMPTEDEROG','');
      SetField(PREFIXE_IMMOCPTE+'_REPRISEDEROG',CAssRepriseDerog(CompteImmo));
      if not VerifCompteAssocie(PC_REPRISEDEROG) then SetField(PREFIXE_IMMOCPTE+'_REPRISEDEROG','');
      SetField(PREFIXE_IMMOCPTE+'_PROVISDEROG',CAssProvisDerog(CompteImmo));
      if not VerifCompteAssocie(PC_PROVISDEROG) then SetField(PREFIXE_IMMOCPTE+'_PROVISDEROG','');
      SetField(PREFIXE_IMMOCPTE+'_DOTATIONEXC',CAssDotationExc (Getfield(PREFIXE_IMMOCPTE+'_COMPTEDOTATION')));
      if not VerifCompteAssocie(PC_DOTATIONEXC) then SetField(PREFIXE_IMMOCPTE+'_DOTATIONEXC','');
      SetField(PREFIXE_IMMOCPTE+'_VACEDEE',CAssVaCedee(CompteImmo));
      if not VerifCompteAssocie(PC_VACEDEE) then SetField(PREFIXE_IMMOCPTE+'_VACEDEE','');
      SetField(PREFIXE_IMMOCPTE+'_AMORTCEDE',GetField(PREFIXE_IMMOCPTE+'_COMPTEAMORT'));
      if not VerifCompteAssocie(PC_AMORTCEDE) then SetField(PREFIXE_IMMOCPTE+'_AMORTCEDE','');
      SetField(PREFIXE_IMMOCPTE+'_VOACEDE',GetField(PREFIXE_IMMOCPTE+'_COMPTEIMMO'));
      if not VerifCompteAssocie(PC_VOACEDE) then SetField(PREFIXE_IMMOCPTE+'_VOACEDE','');
      SetField(PREFIXE_IMMOCPTE+'_REPEXPLOIT',CAssRepExploit (GetfIeld(PREFIXE_IMMOCPTE+'_COMPTEDOTATION')));
      if not VerifCompteAssocie(PC_REPEXPLOIT) then SetField(PREFIXE_IMMOCPTE+'_REPEXPLOIT','');
      SetField(PREFIXE_IMMOCPTE+'_REPEXCEP',CAssRepExc (Getfield(PREFIXE_IMMOCPTE+'_DOTATIONEXC')));
      if not VerifCompteAssocie(PC_REPEXCEP) then SetField(PREFIXE_IMMOCPTE+'_REPEXCEP','');
    end;
  end;
end ;

procedure TOM_IMMOCPTE.OnArgument ( S: String ) ;
begin
  Inherited ;
  {$IFDEF EAGLCLIENT}
  PC_COMPTEIMMO := THEdit(GetControl(PREFIXE_IMMOCPTE+'_COMPTEIMMO'));
  PC_COMPTEAMORT := THEdit(GetControl(PREFIXE_IMMOCPTE+'_COMPTEAMORT'));
  PC_COMPTEDOTATION := THEdit(GetControl(PREFIXE_IMMOCPTE+'_COMPTEDOTATION'));
  PC_COMPTEDEROG := THEdit(GetControl(PREFIXE_IMMOCPTE+'_COMPTEDEROG'));
  PC_REPRISEDEROG := THEdit(GetControl(PREFIXE_IMMOCPTE+'_REPRISEDEROG'));
  PC_DOTATIONEXC := THEdit(GetControl(PREFIXE_IMMOCPTE+'_DOTATIONEXC'));
  PC_PROVISDEROG := THEdit(GetControl(PREFIXE_IMMOCPTE+'_PROVISDEROG'));
  PC_VACEDEE := THEdit(GetControl(PREFIXE_IMMOCPTE+'_VACEDEE'));
  PC_AMORTCEDE := THEdit(GetControl(PREFIXE_IMMOCPTE+'_AMORTCEDE'));
  PC_VOACEDE := THEdit(GetControl(PREFIXE_IMMOCPTE+'_VOACEDE'));
  PC_REPEXCEP := THEdit(GetControl(PREFIXE_IMMOCPTE+'_REPEXCEP'));
  PC_REPEXPLOIT := THEdit(GetControl(PREFIXE_IMMOCPTE+'_REPEXPLOIT'));
  {$ELSE}
  PC_COMPTEIMMO := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_COMPTEIMMO'));
  PC_COMPTEAMORT := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_COMPTEAMORT'));
  PC_COMPTEDOTATION := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_COMPTEDOTATION'));
  PC_COMPTEDEROG := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_COMPTEDEROG'));
  PC_REPRISEDEROG := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_REPRISEDEROG'));
  PC_DOTATIONEXC := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_DOTATIONEXC'));
  PC_PROVISDEROG := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_PROVISDEROG'));
  PC_VACEDEE := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_VACEDEE'));
  PC_AMORTCEDE := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_AMORTCEDE'));
  PC_VOACEDE := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_VOACEDE'));
  PC_REPEXCEP := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_REPEXCEP'));
  PC_REPEXPLOIT := THDBEdit(GetControl(PREFIXE_IMMOCPTE+'_REPEXPLOIT'));
  {$ENDIF}
  PC_COMPTEIMMO.OnElipsisClick := OnElipsisCompteClick;
  PC_COMPTEAMORT.OnElipsisClick := OnElipsisCompteClick;
  PC_COMPTEDOTATION.OnElipsisClick := OnElipsisCompteClick;
  PC_COMPTEDEROG.OnElipsisClick := OnElipsisCompteClick;
  PC_REPRISEDEROG.OnElipsisClick := OnElipsisCompteClick;
  PC_DOTATIONEXC.OnElipsisClick := OnElipsisCompteClick;
  PC_PROVISDEROG.OnElipsisClick := OnElipsisCompteClick;
  PC_VACEDEE.OnElipsisClick := OnElipsisCompteClick;
  PC_AMORTCEDE.OnElipsisClick := OnElipsisCompteClick;
  PC_VOACEDE.OnElipsisClick := OnElipsisCompteClick;
  PC_REPEXCEP.OnElipsisClick := OnElipsisCompteClick;
  PC_REPEXPLOIT.OnElipsisClick := OnElipsisCompteClick;
end ;

procedure TOM_IMMOCPTE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_IMMOCPTE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure RecupereComptesAssocies(QImmo : TQuery; RefCompteImmo : string;var ARecord : TCompteAss);
var QComptesAsso : TQuery; CompteImmo : string;
begin
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
    QComptesAsso := OpenSQL('SELECT * FROM IMMOCPTE WHERE '+PREFIXE_IMMOCPTE+'_COMPTEIMMO = "'+ CompteImmo +'"',TRUE);
    if Not QComptesAsso.EOF then
    begin
      QComptesAsso.First;
      ARecord.Immo := CompteImmo;
      ARecord.Amort := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_COMPTEAMORT').AsString;
      ARecord.Dotation := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_COMPTEDOTATION').AsString;
      ARecord.Derog := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_COMPTEDEROG').AsString;
      ARecord.RepriseDerog := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_REPRISEDEROG').AsString;
      ARecord.DotationExcep := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_DOTATIONEXC').AsString;
      ARecord.ProvisDerog := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_PROVISDEROG').AsString;
      ARecord.VaCedee := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_VACEDEE').AsString;
      ARecord.AmortCede := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_AMORTCEDE').AsString;
      ARecord.VoaCede := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_VOACEDE').AsString;
      ARecord.RepriseExcep := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_REPEXCEP').AsString;
      ARecord.RepriseExploit := QComptesAsso.FindField(''+PREFIXE_IMMOCPTE+'_REPEXPLOIT').AsString;
//EPZ 28/11/00
{ A FAIRE
      ARecord.CreanceCessionActif := QComptesAsso.FindField(PREFIXE_IMMOCPTE+'_CREANCESURCESS').AsString;
      ARecord.ProduitCessionActif := QComptesAsso.FindField(PREFIXE_IMMOCPTE+'_PDTCESSACTIF').AsString;
      ARecord.TVACollectee := QComptesAsso.FindField(PREFIXE_IMMOCPTE+'_TVACOLLECTEE').AsString;
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

procedure TOM_IMMOCPTE.OnElipsisCompteClick(Sender: TObject);
begin
  OnEllipsisCompteImmoGene (Sender);
end;

procedure OnEllipsisCompteImmoGene (Sender : TObject);
var stWhere,TypeCompte : string;
begin
  {$IFDEF EAGLCLIENT}
  TypeCompte := ExtractSuffixe(THEdit(Sender).Name);
  {$ELSE}
  TypeCompte := ExtractSuffixe(THDBEdit(Sender).Name);
  {$ENDIF}
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
  {$IFDEF CCSTD}
  {$IFDEF EAGLCLIENT}
  LookupList(THEdit(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,0)  ;
  {$ELSE}
  LookupList(THDBEdit(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,0)  ;
  {$ENDIF}
  {$ELSE}
  {$IFDEF EAGLCLIENT}
  LookupList(THEdit(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,IMMO_TOM.TAGGENERAUX)  ;
  {$ELSE}
  LookupList(THDBEdit(Sender),TraduireMemoire('Comptes'),'GENERAUX','G_GENERAL','G_LIBELLE',stWhere,'G_GENERAL', True,IMMO_TOM.TAGGENERAUX)  ;
  {$ENDIF}
  {$ENDIF}
end;

Initialization
  registerclasses ( [ TOM_IMMOCPTE ] ) ;
end.
