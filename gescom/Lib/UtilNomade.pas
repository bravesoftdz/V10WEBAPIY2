unit UtilNomade ;

interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,dbtables,
{$ENDIF}
     uToxClasses, EntGc,MenuOLG,
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,ParamSoc,UTob;

function GetNatureNomade(Champ : string; PrefixeAND : boolean = True; VenteAchat : string='VEN') : string;
function VerifEtatBasePCP : integer ;
//procedure GetAchatVente(CodeUser : string = '');
//function GetAchatVente(CodeUser : string) : string;
procedure GetVarPCP(CodeUser : string);
procedure InitMetier;
procedure ScruteTobTox(TobTox : TOB; var MajFiles : boolean);
procedure TrouvePiecesPCP;
procedure TraiteMenusPCP(NumModule : integer);
procedure TestSeriaNomade;

implementation

uses UtilDispGC;

function GetNatureNomade(Champ : string; PrefixeAND : boolean = True; VenteAchat : string='VEN') : string;
var LesNatures, PceVente, PceAchat : string;
begin
  if PrefixeAND then Result:=' AND ' else Result:='';
  if ctxMode in V_PGI.PGIContexte then
    Result:=Result+'('+Champ+' in ("DE","CC"))'
  else
  begin
    LesNatures := VH_GC.PCPPceVente;
    while LesNatures <> '' do
      PceVente := PceVente + ', "'+ReadTokenSt(LesNatures)+'"';
    LesNatures := VH_GC.PCPPceAchat;
    while LesNatures <> '' do
      PceAchat := PceAchat + ', "'+ReadTokenSt(LesNatures)+'"';
    PceVente := Copy(PceVente,2,Length(PceVente));
    PceAchat := Copy(PceAchat,2,Length(PceAchat));
    if (VH_GC.PCPVenteSeria) and (VH_GC.PCPAchatSeria) then
    begin
      if VenteAchat='VEN' then
        Result := Result + '('+Champ+' IN ('+PceVente+'))'
      else if VenteAchat='ACH' then
        Result := Result + '('+Champ+' IN ('+PceAchat+'))'
      else
        Result := Result + '('+Champ+' IN ('+PceVente+', '+PceAchat+'))';
    end else
    if VH_GC.PCPVenteSeria then
    begin
      if (VenteAchat='VEN') or (Venteachat='VENACH') then
        Result := Result + '('+Champ+' IN ('+PceVente+'))'
      else if VenteAchat='ACH' then
        Result := '';
    end else
    if VH_GC.PCPAchatSeria then
    begin
      if (VenteAchat='ACH') or (Venteachat='VENACH') then
        Result := Result + '('+Champ+' IN ('+PceAchat+'))'
      else if VenteAchat='VEN' then
        Result := '';
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret, JTR
Créé le ...... : 03/09/2002
Modifié le ... : 12/12/2003
Description .. : Vérification de l'état de la base PCP :
Suite ........ : Etat
Suite ........ : 1. Vierge : variable TOX $UTIL non existante
Suite ........ : 2. Affectée : aucun article dans la base
Suite ........ : 3. Initialisée : aucune pièce n'a été saisie pour le
Suite ........ : représentant
Suite ........ : 4. Déjà utilisée : il existe au moins une pièce saisie par le
Suite ........ : représentant
Mots clefs ... : PCP;BASE
*****************************************************************}
function VerifEtatBasePCP : integer ;
var QQ : TQuery ;
    NbEnreg : integer ;
    CodeSite : string ;
    LeSite : TCollectionSite ;
begin

  Result := 1; // Vierge : variable TOX $UTIL non existante
  if ( PCP_LesSites <> Nil ) and ( PCP_LesSites.LeSiteLocal <> Nil ) then
  begin
    if PCP_LesSites.LeSiteLocal.LesVariables.Find ( '$UTIL' ) <> nil then
    begin
      QQ := nil ;
      if ( V_PGI.User <> 'CEG' ) and (not V_PGI.Superviseur)  then
      begin
        CodeSite := '' ;
        if PCP_LesSites <> nil then
        begin
          LeSite := PCP_LesSites.FindVariableValue( '$UTIL', V_PGI.User, '' ) ;
          if LeSite <> nil then
            CodeSite := LeSite.SSI_CODESITE ;
        end ;
        if CodeSite <> '' then
        QQ := OpenSQL ( 'select count(*) from PIECE where ' + GetNatureNomade ( 'GP_NATUREPIECEG', False, 'VENACH' ) +
            ' and GP_SOUCHE in (select GSP_SOUCHE from SITEPIECE where GSP_CODESITE="' + CodeSite + '")', True ) ;
      end else
      begin
        QQ := OpenSQL ( 'select count(*) from PIECE where ' + GetNatureNomade ( 'GP_NATUREPIECEG', False, 'VENACH' ) +
            ' and GP_SOUCHE in (select GSP_SOUCHE from SITEPIECE)', True ) ;
      end ;
      NbEnreg := 0 ;
      if QQ <> nil then
      begin
        if not QQ.EOF then NbEnreg := QQ.Fields[0].AsInteger ;
        Ferme ( QQ ) ;
      end ;
      // Déjà utilisée : il existe au moins une pièce saisie par le représentant
      if NbEnreg > 0 then BEGIN Result := 4 ; exit END ;

      QQ := OpenSQL ( 'select count(*) from ARTICLE', True ) ;
      if not QQ.EOF then NbEnreg := QQ.Fields[0].AsInteger else NbEnreg := 0 ;
      Ferme ( QQ ) ;
      if NbEnreg > 0 then Result := 3 // Initialisée : aucune pièce n'a été saisie pour le représentant
      else Result := 2 ; // Affectée : aucun article dans la base
    end ;
  end ;
end ;

procedure GetVarPCP(CodeUser : string);
var Req : string;
    Qte : integer;
    Qry : TQuery;
begin
  InitLaVariableGCPCP;
  VH_GC.PCPPrefixe := CodeUser;
  TrouvePiecesPCP;
  // Recherche représentant lié
  Req := 'FROM COMMERCIAL WHERE GCL_UTILASSOCIE="'+CodeUser+'" AND GCL_TYPECOMMERCIAL="REP"';
  Qry := OpenSQL('SELECT COUNT(*) AS QTE '+Req,True);
  if Not Qry.Eof then
    Qte := Qry.Fields[0].asInteger
    else
    Qte := 0;
  Ferme(Qry);
  if Qte <> 1 then
  begin
    VH_GC.PCPRepresentant := '';
  end else
  begin
    Qry := OpenSQL('SELECT GCL_COMMERCIAL '+Req,True);
    VH_GC.PCPRepresentant := Qry.Fields[0].asString;
  end;
  Ferme(Qry);
  // Recherche Vente/Achat de l'utilisateur
  Qry := OpenSQL('SELECT CPU_PCPACHAT, CPU_PCPVENTE FROM CPPROFILUSERC '+
                 'WHERE CPU_USER ="'+CodeUser+'"', True);
  if not Qry.EOF then
  begin
    VH_GC.PCPUsVte := (Qry.FindField('CPU_PCPVENTE').AsString = 'X');
    VH_GC.PCPUsAch := (Qry.FindField('CPU_PCPACHAT').AsString = 'X');
  end;
  Ferme(Qry);
end;

//procedure GetAchatVente(CodeUser : string = '');
(*function GetAchatVente(CodeUser : string) : string;
var Qry : TQuery;
    UsVente, UsAchat : boolean;
begin
  UsVente := false;
  UsAchat := false;
  Qry := OpenSQL('SELECT CPU_PCPACHAT, CPU_PCPVENTE FROM CPPROFILUSERC '+
                 'WHERE CPU_USER ="'+CodeUser+'"', True);
  if not Qry.EOF then
  begin
    UsVente := (Qry.FindField('CPU_PCPVENTE').AsString = 'X');
    UsAchat := (Qry.FindField('CPU_PCPACHAT').AsString = 'X');
  end else
  begin
    UsVente := false;
    UsAchat := false;
  end;
  Ferme(Qry);
  if (UsVente) and (UsAchat) then
    Result := 'VEN;ACH'
  else if UsVente then
    Result := 'VEN'
  else if UsAchat then
    Result := 'ACH';

{  if V_PGI.User <> 'CEG' then
  begin
    if UsVente then
    begin
      if (not VH_GC.PCPVenteSeria) and (V_PGI.VersionDemo) then
        VH_GC.PCPVenteSeria := True;
    end else
      VH_GC.PCPVenteSeria := False;
    if UsAchat then
    begin
      if (not VH_GC.PCPAchatSeria) and (V_PGI.VersionDemo) then
        VH_GC.PCPAchatSeria := True;
    end else
      VH_GC.PCPAchatSeria := False;
  end;
}
{
  if (VH_GC.PCPVenteSeria) and (UsVente) then
    VH_GC.PCPVenteSeria := True
    else
    VH_GC.PCPVenteSeria := False;
  if (VH_GC.PCPAchatSeria) and (UsAchat) then
    VH_GC.PCPAchatSeria := True
    else
    VH_GC.PCPAchatSeria := False;
}
end;
*)
procedure InitMetier;
begin
  if GetParamSoc('SO_METIER') = 'MOD' then
  begin
    if not (ctxMode in V_PGI.PGIContexte) then
      V_PGI.PGIContexte := V_PGI.PGIContexte + [ctxMode];
  end;
  if ctxMode in V_PGI.PGIContexte then
    Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CMBOS5.HLP'
    else
{$IFDEF BTP}
    Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CBPOS3.HLP';
{$ELSE}
    Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CGS5.HLP';
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 22/12/2003
Modifié le ... :
Description .. : Traitement des données traitées par la Tox
Suite ........ : Exécuté après traitement de la Tox
Mots clefs ... : TOX;PCP
*****************************************************************}
procedure ScruteTobTox(TobTox : TOB; var MajFiles : boolean);
var Cpt, Cpt1 : integer;
    TobToxCondition, TobToxTable, TobPourMaj : TOB;

    procedure MajDataTox(NomTable : string);
    var TobTmp, TobTmp1, TobAdupliquer : TOB;
    begin
      if NomTable = 'PIECE' then
      begin
        TobTmp := TobToxTable.FindFirst(['GP_VIVANTE'],['X'],True);
        while TobTmp <> nil do
        begin
          if GetInfoParpiece(TobTmp.GetValue('GP_NATUREPIECEG'),'GPP_ACTIONFINI') = 'TOX' then
          begin
            TobADupliquer := TOB.Create('PIECE',TobPourMaj,-1);
            TobADupliquer.Dupliquer(TobTmp,true,true,true);
            TobADupliquer.PutValue('GP_VIVANTE','-');
            TobTmp1 := TobADupliquer.FindFirst(['GL_VIVANTE'],['X'],True);
            while TobTmp1 <> nil do
            begin
              TobTmp1.PutValue('GL_VIVANTE','-');
              TobTmp1 := TobADupliquer.FindNext(['GL_VIVANTE'],['X'],True);
            end;
          end;
          TobTmp := TobToxTable.FindNext(['GP_VIVANTE'],['X'],True);
        end;
      end;
    end;

begin
  TobPourMaj := TOB.Create('LES MAJ',nil,-1);
  if (TobTox.FieldExists('EVENEMENT')) then
  begin
    for Cpt := 0 to TobTox.Detail.Count-1 do
    begin
      TobToxCondition := TobTox.Detail[Cpt];
      if TobToxCondition.FieldExists('CONDITION') then
      begin
        for Cpt1 := 0 to TobToxCondition.Detail.Count-1 do
        begin
          TobToxTable := TobToxCondition.Detail[Cpt1];
          if TobToxTable.FieldExists('TOX_TABLE') then
            MajDataTox(TobToxTable.GetValue('TOX_TABLE'));
        end;
      end;
    end;
  end;
  if TobPourMaj.Detail.count > 0 then
    TobPourMaj.UpdateDB(False);
  FreeAndNil(TobPourMaj);
end;

{***********A.G.L.***********************************************
Auteur  ...... : JTR
Créé le ...... : 19/03/2004
Modifié le ... :
Description .. : Mémorise les pièces utilisables dans PCP
Mots clefs ... : TOX;PCP
*****************************************************************}
procedure TrouvePiecesPCP;
var Cpt : integer;
begin
  VH_GC.PCPPceVente := '';
  VH_GC.PCPPceAchat := '';
  for Cpt := 0 to VH_GC.TOBParPiece.detail.count -1 do
  begin
    if VH_GC.TOBParPiece.detail[Cpt].GetValue('GPP_MASQUERNATURE') = '-' then
    begin
      if VH_GC.TOBParPiece.detail[Cpt].GetValue('GPP_TRSFVENTE') = 'X' then
        VH_GC.PCPPceVente := VH_GC.PCPPceVente + ';' + VH_GC.TOBParPiece.detail[Cpt].GetValue('GPP_NATUREPIECEG')
      else if VH_GC.TOBParPiece.detail[Cpt].GetValue('GPP_TRSFACHAT') = 'X' then
        VH_GC.PCPPceAchat := VH_GC.PCPPceAchat + ';' + VH_GC.TOBParPiece.detail[Cpt].GetValue('GPP_NATUREPIECEG');
    end;
  end;
  VH_GC.PCPPceVente := Copy(VH_GC.PCPPceVente,2,length(VH_GC.PCPPceVente))+';';
  VH_GC.PCPPceAchat := Copy(VH_GC.PCPPceAchat,2,length(VH_GC.PCPPceAchat))+';';
end;

procedure TraiteMenusPCP(NumModule : integer);
begin
  if (VH_GC.PCPPceVente = '') and (VH_GC.PCPPceAchat = '') then exit;

  if NumModule = 285 then
  begin
    // Devis
    if pos('DBT;',VH_GC.PCPPceVente) = 0 then
    begin
     FMenuG.RemoveItem(30201); // Saisie
     FMenuG.RemoveItem(30601); // Modification
    end;
  end else
  if NumModule = 149 then // fichier de base
  begin
    FMenuG.RemoveItem(30101); // assistant de creation de tarif
    FMenuG.RemoveItem(149292); // supression des articles
    FMenuG.RemoveItem(149295); // mise ajour tarifaire
    FMenuG.RemoveItem(32601);  // dépots
    FMenuG.RemoveItem(32602); // emplacements
    FMenuG.RemoveItem(33105); // Cub decisionnel vente
    FMenuG.RemoveItem(30506); // Activation client
    FMenuG.RemoveItem(30507); // Supression Client

    FMenuG.RemoveItem(31401); // Consultation des tarifs fournisseurs
    FMenuG.RemoveItem(31403); // Mise a jour
    FMenuG.RemoveItem(31404); // import catalogue

    FMenuG.RemoveItem(31704); // modification en serie
    FMenuG.RemoveItem(31705); // mise en sommeil
    FMenuG.RemoveItem(31706); // activation
    FMenuG.RemoveItem(31707); // Suppression

    FMenuG.RemoveItem(30503); // Mise a jour des commerciaux

  	FmenuG.RemoveGroup (-147700,true); // Stocks parametrage
  	FmenuG.RemoveGroup (-149400,True); // Menu fournisseurs
  	FmenuG.RemoveGroup (-149900,True); // Menu des ressources
  	FmenuG.RemoveGroup (-149600,True); // Menu des utilitaires
  end;
  if NumModule = 148 then // fichier de base
  begin
  	FmenuG.RemoveGroup (148100,True); // Menu des utilitaires
  end;
end;

procedure TestSeriaNomade;
begin
  if not V_PGI.VersionDemo then exit;
{$IFNDEF BTP}
  if ((VH_GC.PCPUsVte) and (not VH_GC.PCPVenteSeria)) and ((VH_GC.PCPUsAch) and (not VH_GC.PCPAchatSeria)) then
  begin
    SetLength(SeriaCode,2);
    SetLength(SeriaTitre,2);
    SeriaCode[0] := '00262060';
    SeriaTitre[0] := 'Devis Portable';
    SeriaCode[1] := '00251'+GCVersionSeria;
    SeriaTitre[1] := 'Saisie NOMADE Achats';
  end else if (VH_GC.PCPUsVte) and (not VH_GC.PCPVenteSeria) then
  begin
    SetLength(SeriaCode,1);
    SetLength(SeriaTitre,1);
    SeriaCode[0]  := '00262060';
    SeriaTitre[0] := 'Devis Portable';
  end else if (VH_GC.PCPUsAch) and (not VH_GC.PCPAchatSeria) then
  begin
    SetLength(SeriaCode,1);
    SetLength(SeriaTitre,1);
    SeriaCode[0]  := '00251'+GCVersionSeria;
    SeriaTitre[0] := 'Saisie NOMADE Achats';
  end;
{$ENDIF}
  FMenuG.SetSeria(GCCodeDomaine, GCCodesSeria, GCTitresSeria);
end;
end.

