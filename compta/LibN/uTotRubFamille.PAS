{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/04/2001
Modifié le ... : 25/08/2005 - GCO - FQ 16417
Description .. : Source TOT de la TABLE : TTRUBFAMILLE (TTRUBFAMILLE)
Mots clefs ... : TOF;TTRUBFAMILLE
*****************************************************************}
{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
                 18/04/01    ??     Création de l'unité
 4.3.1.101.001   26/05/03    JP     Reprise de la TOT pour eCCS5
 4.4.0.101.001   12/09/03    JP     Gestion du bouton complément (Valable à partir de l'AGL 560n)
--------------------------------------------------------------------------------------}
Unit UTOTRUBFAMILLE ;

Interface

uses
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  StdCtrls,
  Controls,
  Classes,
  Forms,
  Sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOT ,
  HTB97,
  Ent1,
  UTOB,
  FamRub_tof,
  UtilSoc,    // SupprimeCycleRevisionDesGene
  GalOutil    // NoDossierBaseCommune
  ;

////////////////////////////////////////////////////////////////////////////////

Type
  TOT_CPFAMILLE = Class ( TOT )
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnUpdateRecord           ; override ;
    procedure OnAfterUpdateRecord      ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument  (S : String ); override ;
    procedure OnComplement(Prefixe, Tipe, Code : string); override ;
  private
    { Déclarations privées }
  end ;

////////////////////////////////////////////////////////////////////////////////

Type
  TOT_CPRUBFAMILLE = Class ( TOT_CPFAMILLE )
    procedure OnArgument  (S : String ); override ;
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnComplement(Prefixe, Tipe, Code : string); override ;
  private
    { Déclarations privées }
    function  SupprimeFamille : Boolean ;
  end ;

////////////////////////////////////////////////////////////////////////////////

Type
  TOT_CPPLANREVISION = Class ( TOT_CPFAMILLE )
    procedure OnArgument  (S : String ); override ;
    procedure OnNewRecord              ; override ;
    procedure OnDeleteRecord           ; override ;
    procedure OnComplement(Prefixe, Tipe, Code : string); override ;
  private
    { Déclarations privées }
    function SupprimePlanRevision : Boolean;
  end ;

////////////////////////////////////////////////////////////////////////////////

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  ParamSoc;   // GetParamSocSecur, SetParamSoc


const
  HMess : array [0..3] of string = (
        '0;?caption?;Cette famille ne peut pas être supprimée, elle est référencée par une rubrique.;W;O;O;O;',
        '1;?caption?;Il n''existe aucune rubrique, vous ne pouvez pas activer cette fonction;W;O;O;O;',
        '2;?caption?;Ce plan de révision ne peut pas être supprimé, il est référencé par un cycle de révision.;W;O;O;O;',
        '3;?caption?;Il n''existe aucun cycle de révision, vous ne pouvez pas activer cette fonction;W;O;O;O;');

{---------------------------------------------------------------------------------------}
procedure TOT_CPFAMILLE.OnNewRecord ;
{---------------------------------------------------------------------------------------}
var bStdCegid : boolean;
begin
  Inherited ;

  if ctxStandard in V_PGI.PGIContexte then
  begin
    bStdCegid := EstSpecif('51502');
    if (bStdCegid) then
    begin
      SetField('YDS_PREDEFINI','CEG');
      SetField('YDS_CODE','@');
    end
    else begin
      SetField('YDS_PREDEFINI','STD');
      SetField('YDS_CODE','$');
    end;
    // GCO - 25/08/2005 - FQ 16417
    SetField('YDS_NODOSSIER', NoDossierBaseCommune);
  end
  else
  begin
    SetField('YDS_PREDEFINI','DOS');
    SetField('YDS_NODOSSIER',V_PGI.NoDossier);
  end;
end ;

{---------------------------------------------------------------------------------------}
procedure TOT_CPFAMILLE.OnDeleteRecord ;
{---------------------------------------------------------------------------------------}
var
  bStdCegid : boolean;
  stCode : string;
begin
  stCode := GetField('YDS_CODE');
  if (stCode <> '' ) then
  begin
    bStdCegid := EstSpecif('51502');
    if ((bStdCegid) and (ctxStandard in V_PGI.PGIContexte)) then bStdCegid :=True
    else bStdCegid := False;
    if (((ctxStandard in V_PGI.PGIContexte) and (not (bStdCegid))) and (stCode[1]='@')) then
    begin
      LastError := -1;
      LastErrorMsg := 'Suppression d''un standard Cegid impossible';
    end else
    if (( (not (ctxStandard in V_PGI.PGIContexte)) and (not bStdCegid)) and ((stCode[1]='@') or (stCode[1]='$'))) then
    begin
      LastError := -2;
      LastErrorMsg := 'Suppression d''un standard impossible';
    end;
  end;
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOT_CPFAMILLE.OnUpdateRecord ;
{---------------------------------------------------------------------------------------}
var
  stCode : string;
  bStdCegid : boolean;
begin
  if (not (DS.state in [dsEdit,dsInsert])) then exit;

  stCode := GetField('YDS_CODE');
  if stCode = '' then Exit;

  (*
  bStdCegid := EstSpecif('51502');

  if ((bStdCegid) and (ctxStandard in V_PGI.PGIContexte)) then
    bStdCegid :=True
  else
    bStdCegid := False;


  if (((ctxStandard in V_PGI.PGIContexte) and (stCode[1]<>'$')) or
    ((not (ctxStandard in V_PGI.PGIContexte)) and ((stCode[1]='$') or (stCode[1]='@')))) then
  begin
    if (not bStdCegid) then
    begin
      LastError := -1;
      LastErrorMsg := 'Code famille incorrect.'+#10#13+'Les codes commençant par "@" sont réservés à Cegid.'+
            #10#13+'Les codes commençant par "$" sont réservés aux standards.';
    end else if stCode[1]<>'@' then SetField('YDS_CODE','@'+GetField('YDS_CODE'));
  end;*)

  // GCO - 31/08/2004
  if (ctxStandard in V_PGI.PGIContexte) and EstSpecif('51502') then
    bStdCegid := True
  else
    bStdCegid := False;

  // Lancement par le CCSTD
  if (CtxStandard in V_PGI.PGIContexte) then
  begin
    // Si je peux créer des STANDARD CEGID
    if bStdCegid then
    begin
      if not (stCode[1]='@') then
        SetField('YDS_CODE','@'+GetField('YDS_CODE'));
    end
    else
    begin // Config CLIENT
      if not (stCode[1]='$') then
        SetField('YDS_CODE','$'+GetField('YDS_CODE'));
    end;
  end;

  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOT_CPFAMILLE.OnAfterUpdateRecord ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOT_CPFAMILLE.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOT_CPFAMILLE.OnArgument(S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;

  OkComplement := True;
  bComplement.GlobalIndexImage := 'Z0008_S16G1';
end ;

{Appel de l'écran de sélection des familles de rubriques : FAMRUB_TOF
{---------------------------------------------------------------------------------------}
procedure TOT_CPFAMILLE.OnComplement(Prefixe, Tipe, Code : string) ;
{---------------------------------------------------------------------------------------}
begin
  inherited ;

end ;
////////////////////////////////////////////////////////////////////////////////










////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOT_CPRUBFAMILLE.OnArgument(S : String ) ;
begin
  inherited;
  bComplement.Hint := TraduireMemoire('Affectation des rubriques aux familles');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOT_CPRUBFAMILLE.OnNewRecord;
begin
  inherited;
  SetField('YDS_TYPE','RBF');
  SetField('YDS_LIBRE','');
end;

{On s'assure que la famille n'est pas rattachée à une rubrique avant de la supprimer
{---------------------------------------------------------------------------------------}
function TOT_CPRUBFAMILLE.SupprimeFamille : Boolean ;
{---------------------------------------------------------------------------------------}
var
  Qr : TQuery ;
  Tb : TOB;
  n  : Integer;
  Trouver   : Boolean ;
  St, StLib : string ;
begin
  Result:=True ;
  Trouver:=False ;

  Qr := OpenSql('Select RB_FAMILLES from RUBRIQUE',True) ;
  Tb := Tob.Create('', nil, -1);

  if not Qr.EOF then begin
    Tb.LoadDetailDB('', '', '', Qr, False);
    for n := 0 to Tb.Detail.Count - 1 do begin
      if Trouver then Break;
      StLib := Tb.Detail[n].GetValue('RB_FAMILLES');
       while StLib <> '' do begin
         St := ReadTokenSt(StLib) ;
         if St = GetField('YDS_CODE') then begin
           Trouver := True ;
           Break ;
         end;
      end;
    end;
  end;

  if Trouver then begin
    if HShowMessage(HMess[0], '', '') <> mrYes then
      Result := False ;
  end;

  Ferme(Qr) ;
  FreeAndNil(Tb);
  Screen.Cursor:=SyncrDefault ;
end;



////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOT_CPRUBFAMILLE.OnDeleteRecord ;
begin
  if not SupprimeFamille then
  begin
    LastError := -3;
    Exit;
  end;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
// Appel de l'écran de sélection des familles de rubriques : FAMRUB_TOF
////////////////////////////////////////////////////////////////////////////////
{---------------------------------------------------------------------------------------}
procedure TOT_CPRUBFAMILLE.OnComplement(Prefixe, Tipe, Code : string) ;
{---------------------------------------------------------------------------------------}
Var Famille,LibFam : String ;
    Q : Tquery ;
    Vide : Boolean ;
begin
  inherited ;
  Q := OpenSql('Select * from Rubrique WHERE RB_CLASSERUB NOT IN("CDR","TRE")',True);

  Vide := Q.Eof ;
  Ferme(Q) ;

  if Vide then
  begin
    HShowMessage(HMess[1],'','');
    Exit;
  end;

  Famille := '';
  LibFam  := '';
  ParametrageFamilleRubrique('','',Famille,LibFam,False);
  Screen.Cursor := SyncrDefault;
end ;
////////////////////////////////////////////////////////////////////////////////


















////////////////////////////////////////////////////////////////////////////////
procedure TOT_CPPLANREVISION.OnArgument(S : String ) ;
begin
  inherited;
  OkComplement := False;
  bComplement.Hint := TraduireMemoire('Affectation des cycles de révision aux plans de révision');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOT_CPPLANREVISION.OnNewRecord;
begin
  inherited;
  SetField('YDS_TYPE', 'PDR');
  SetField('YDS_LIBRE', Now );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2005
Modifié le ... : 25/08/2005
Description .. :
Mots clefs ... :
*****************************************************************}
function TOT_CPPLANREVISION.SupprimePlanRevision : Boolean;
begin
  Result  := False ;

  if ExisteSQL('SELECT RB_RUBRIQUE FROM RUBRIQUE WHERE ' +
               'RB_NATRUB = "CPT" AND RB_CLASSERUB = "CDR" AND ' +
               'RB_FAMILLES LIKE "%' + GetField('YDS_CODE') + ';%" AND ' +
               '((RB_PREDEFINI = "CEG") OR (RB_PREDEFINI = "STD"))') then
    HShowMessage(HMess[2], '', '')
  else
    Result := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/02/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOT_CPPLANREVISION.OnDeleteRecord ;
begin
  if not SupprimePlanRevision then
  begin
    LastError := -3;
    Exit;
  end;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////
// Appel de l'écran de sélection des cycles de révision
////////////////////////////////////////////////////////////////////////////////
{---------------------------------------------------------------------------------------}
procedure TOT_CPPLANREVISION.OnComplement(Prefixe, Tipe, Code : string) ;
{---------------------------------------------------------------------------------------}
var Famille,LibFam : String ;
    Q : Tquery ;
    Vide : Boolean ;
    lStWhere : string;
begin
  inherited ;
  if EstSpecif('51502') and (ctxStandard in V_PGI.PGIContexte) then
    lStWhere := 'RB_PREDEFINI = "CEG"'
  else
    lStWhere := 'RB_PREDEFINI = "STD"';

  Q := OpenSql('SELECT * FROM RUBRIQUE WHERE RB_CLASSERUB = "CDR" AND ' + lStwhere,True);
  Vide := Q.Eof ;
  Ferme(Q) ;

  if Vide then
  begin
    HShowMessage(HMess[3],'','');
    Exit;
  end;

  Famille := '';
  LibFam  := '';
  ParametrageFamilleRubrique('','',Famille,LibFam,False);
  Screen.Cursor := SyncrDefault;
end ;

////////////////////////////////////////////////////////////////////////////////


Initialization
  registerclasses ( [ TOT_CPRUBFAMILLE ] );
  registerclasses ( [ TOT_CPPLANREVISION ] );

end.


