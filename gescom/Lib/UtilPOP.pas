unit UtilPOP ;

interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,dbtables,
{$ENDIF}
     uToxClasses, EntGc,MenuOLG,
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,ParamSoc,UTob;

function GetNaturePOP(Champ : string; PrefixeAND : boolean = True) : string;
function VerifEtatBasePCP : integer ;
//function GetEnvPop : String;
procedure GetAchatVente(CodeUser : string = '');
function ActiveEvent(const codeevent: string): boolean;
procedure InitMetier;
procedure ScruteTobTox(TobTox : TOB; var MajFiles : boolean);
procedure TrouvePiecesPCP;
procedure TraiteMenusPCP(NumModule : integer);

implementation

function GetNaturePOP(Champ : string; PrefixeAND : boolean = True) : string;
begin
  if PrefixeAND then Result:='and ' else Result:='';
{$IFDEF BTP}
    Result:=Result+'('+Champ+' in ("DBT","ETU"))'
{$ELSE}
  if ctxMode in V_PGI.PGIContexte then
    Result:=Result+'('+Champ+' in ("DE","CC"))'
    else
    Result:=Result+'('+Champ+' in ("DE","CC","CCE"))';
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 03/09/2002
Modifié le ... :   /  /
Description .. : Vérification de l'état de la base PCP :
Suite ........ : Etat
Suite ........ : 1. Vierge : variable TOX $REP non existante
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
    CodeRep,CodeSite : string ;
    LeSite : TCollectionSite ;
begin

  if (V_PGI.User = 'CEG') then // Connexion CEGID sur une base non vierge
  begin
    QQ := OpenSQL ( 'select count(*) from UTILISAT where US_UTILISATEUR<>"'+V_PGI.User+'"', True ) ;
    if not QQ.EOF then NbEnreg := QQ.Fields[0].AsInteger else NbEnreg := 0 ;
    Ferme ( QQ ) ;
    // Déjà utilisée : il existe un utilisateur différent de CEGID
    if NbEnreg > 0 then begin Result := 5 ; exit end ;
  end ;

  Result := 1; // Vierge : variable TOX $REP non existante
  //if ExisteSQL('select SVA_NOMVARIABLE from STOXVARS where SVA_NOMVARIABLE="$REP"') then
  //begin
  if ( PCP_LesSites <> Nil ) and ( PCP_LesSites.LeSiteLocal <> Nil ) then
  begin
    if PCP_LesSites.LeSiteLocal.LesVariables.Find ( '$REP' ) <> nil then
    begin
      QQ := nil ;
      if ( V_PGI.User <> 'CEG' ) then
      begin
        CodeRep := GetColonneSQL ( 'COMMERCIAL', 'GCL_COMMERCIAL', 'GCL_TYPECOMMERCIAL="REP" and GCL_UTILASSOCIE="' + V_PGI.User + '"' ) ;
        CodeSite := '' ;
        if PCP_LesSites <> nil then
        begin
          LeSite := PCP_LesSites.FindVariableValue( '$REP', CodeRep, '' ) ;
          if LeSite <> nil then CodeSite := LeSite.SSI_CODESITE ;
        end ;
        if ( CodeRep <> '' ) and ( CodeSite <> '' ) then
        QQ := OpenSQL ( 'select count(*) from PIECE where ' + GetNaturePOP ( 'GP_NATUREPIECEG', False ) +
            ' and GP_SOUCHE in (select GSP_SOUCHE from SITEPIECE where GSP_CODESITE="' + CodeSite + '")', True ) ;
      end else
      begin
        QQ := OpenSQL ( 'select count(*) from PIECE where ' + GetNaturePOP ( 'GP_NATUREPIECEG', False ) +
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
{
function GetEnvPop : String;
//var Buffer: array[0..1023] of Char;
//    sWinPath,sIni : String;
var sIni : String;
begin
Result:='0';
// ---- Répertoires
//GetWindowsDirectory(Buffer,1023);
//SetString(sWinPath, Buffer, StrLen(Buffer));

// ---- Fichier ini
//sIni:=sWinPath+'\CEGIDPOP.INI';
// Répertoire par défaut = répertoire de l'exe
sIni:='CEGIDPCP.INI';
if FileExists(sIni) then Result:='1';
end;
}

procedure GetAchatVente(CodeUser : string = '');
var Qry : TQuery;
    UsVente, UsAchat : boolean;
    User, Req : string;
    Qte : integer;
begin
  if CodeUser = '' then
    User := V_PGI.User
    else
    User := CodeUser;
  VH_GC.PCPPrefixe := User;
  //Calcul VH_GC.PCPRepresentant (représentant associé à l'utilisateur)
  Req := 'FROM COMMERCIAL WHERE GCL_UTILASSOCIE="'+V_PGI.User+'" AND GCL_TYPECOMMERCIAL="REP"';
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
  InitLaVariableGCPCP;
  Qry := OpenSQL('SELECT CPU_PCPACHAT, CPU_PCPVENTE FROM CPPROFILUSERC '+
                 'WHERE CPU_USER ="'+User+'"', True);
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
  if (VH_GC.PCPVenteSeria) and (UsVente) then
    VH_GC.PCPVenteSeria := True
    else
    VH_GC.PCPVenteSeria := False;
  if (VH_GC.PCPAchatSeria) and (UsAchat) then
    VH_GC.PCPAchatSeria := True                                           
    else
    VH_GC.PCPAchatSeria := False;
end;

function ActiveEvent(const CodeEvent: string): boolean;
var Evt : TCollectionEvent;
begin
  try
    Evt := TX_LesEvenements.Find(CodeEvent);
    if assigned(Evt) then
    begin
      if not Evt.SEV_ACTIF then
      begin
        Evt.SEV_ACTIF := True;
        result := Evt.Update();
      end else
        result := true;
    end else
      result := False;
  except
    result := false;
  end;
end;

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
    Application.HelpFile := ExtractFilePath(Application.ExeName) + 'CGS5.HLP';
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
    if VH_GC.TOBParPiece.detail[Cpt].GetValue('GPP_TRSFVENTE') = 'X' then
      VH_GC.PCPPceVente := VH_GC.PCPPceVente + ';' + VH_GC.TOBParPiece.detail[Cpt].GetValue('GPP_NATUREPIECEG')
      else
      VH_GC.PCPPceAchat := VH_GC.PCPPceAchat + ';' + VH_GC.TOBParPiece.detail[Cpt].GetValue('GPP_NATUREPIECEG');
  end;
  VH_GC.PCPPceVente := Copy(VH_GC.PCPPceVente,2,length(VH_GC.PCPPceVente))+';';
  VH_GC.PCPPceAchat := Copy(VH_GC.PCPPceAchat,2,length(VH_GC.PCPPceAchat))+';';
end;

procedure TraiteMenusPCP(NumModule : integer);
begin
  if (VH_GC.PCPPceVente = '') and (VH_GC.PCPPceAchat = '') then exit;
//VENTE
  if NumModule = 279 then
  begin
    // Devis
    if pos('DE;',VH_GC.PCPPceVente) = 0 then
    begin
     FMenuG.RemoveItem(30201); // Saisie
     FMenuG.RemoveItem(30221); // Modification
     FMenuG.RemoveItem(30231); // Duplication
    end;
    // Commande
    if pos('CC;',VH_GC.PCPPceVente) = 0 then
    begin
     FMenuG.RemoveItem(30241);   // Saisie
     FMenuG.RemoveItem(30223);   // Modification
     FMenuG.RemoveItem(30233);   // Duplication
     FMenug.RemoveGroup(279300); // Génération
    end;
    // Commande échantillon
    if pos('CCE;',VH_GC.PCPPceVente) = 0 then
    begin
     FMenuG.RemoveItem(30242);  // Saisie
     FMenuG.RemoveItem(30226);  // Modification
     FMenuG.RemoveItem(279133); // Duplication
    end;
    // Proforma
    if pos('PRO;',VH_GC.PCPPceVente) = 0 then
    begin
     FMenuG.RemoveItem(30202);  // Saisie
     FMenuG.RemoveItem(30222); // Modification
     FMenuG.RemoveItem(30232); // Duplication
    end;
    // Commande + Commande échantillon
    if (pos('CC;',VH_GC.PCPPceVente) = 0) and (pos('CCE',VH_GC.PCPPceVente) = 0) then
     FMenuG.RemoveItem(-279110);
  end else
//ACHAT
  if NumModule = 280 then
  begin
    // Proposition d'achat
    if pos('DEF;',VH_GC.PCPPceAchat) = 0 then
    begin
     FMenuG.RemoveItem(31206); //Saisie
     FMenuG.RemoveItem(31213); //Modification
     FMenuG.RemoveItem(31224); //Duplication
    end;
    // Commande
    if pos('CF;',VH_GC.PCPPceAchat) = 0 then
    begin
     FMenuG.RemoveItem(31201); //Saisie
     FMenuG.RemoveItem(31211); //Modification
     FMenuG.RemoveItem(31221); //Duplication
     FMenug.RemoveGroup(280300); // Génération
    end;
  end;
end;

end.

