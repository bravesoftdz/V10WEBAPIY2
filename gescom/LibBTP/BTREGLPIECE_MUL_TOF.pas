{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREGLPIECE_MUL ()
Mots clefs ... : TOF;BTREGLPIECE_MUL
*****************************************************************}
Unit BTREGLPIECE_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Ent1,
     HTB97,
     HPanel,
     windows,
     messages,
{$IFDEF EAGLCLIENT}
      eMul,
      Maineagl,
      HQry,
{$ELSE}
      db,
      {$IFNDEF DBXPRESS}
      dbTables,
      {$ELSE}
      uDbxDataSet,
      {$ENDIF}
      DBGrids,
      mul,
      FE_Main,
{$ENDIF}
      forms,
      sysutils,
      HDB,
      ParamSoc,
      utob,
      ComCtrls,
      HCtrls,
      HEnt1,
      HMsgBox,
      UTOF,
      ConfidentAffaire,
      M3FP,
      AglInitBtp,
      uTofAfBaseCodeAffaire,
      UEntCommun;

Type
  TOF_BTREGLPIECE_MUL = Class (TOF_AFBASECODEAFFAIRE)
  private
    Fliste  : THDbGrid;
    TypeReglt : String;
    //
    procedure OnDblClick(Sender: Tobject);
    //
    procedure DemandeReglementSitDAC;
    procedure SaisieReglementsPiece(Cledoc : R_Cledoc; Tiers : string);
    procedure ControleChamp(Champ, Valeur: String);

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire( Var Aff,Aff0,Aff1,Aff2,Aff3,Aff4,Aff_,Aff0_,Aff1_,Aff2_,Aff3_,Aff4_,Tiers,Tiers_ : THEdit ) ; override ;

  end ;


Implementation
uses FactUtil;

procedure TOF_BTREGLPIECE_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLPIECE_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLPIECE_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLPIECE_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLPIECE_MUL.OnArgument (S : String ) ;
Var CC 			: THValComboBox ;
    Critere	: string;
    Champ		: string;
    Valeur	: string;
    i_ind 	: integer;
begin
	fMulDeTraitement := true;
  Inherited ;
  fTableName := 'PIECE';
	Critere:=(Trim(ReadTokenSt(S)));

  While (Critere <> '') do
  BEGIN
    i_ind:=pos(':',Critere);
    if i_ind = 0 then i_ind:=pos('=',Critere);
    if i_ind <> 0 then
       begin
       Champ:=copy(Critere,1,i_ind-1);
       Valeur:=Copy (Critere,i_ind+1,length(Critere)-i_ind);
       end
    else
       Champ := Critere;
    Controlechamp(Champ, Valeur);
    Critere:=(Trim(ReadTokenSt(S)));
  END;

  CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ;
  if CC<>Nil then PositionneEtabUser(CC) ;

  CC:=THValComboBox(GetControl('GP_DOMAINE')) ;
  if CC<>Nil then PositionneDomaineUser(CC) ;

  if Assigned(GetControl('Fliste')) then
  begin
    Fliste := THDBGrid(ecran.FindComponent('Fliste'));
    Fliste.OnDblClick := OnDblClick;
  end;

end ;

procedure TOF_BTREGLPIECE_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLPIECE_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLPIECE_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTREGLPIECE_MUL.OnDblClick(Sender : TOBJect);
begin

  DemandeReglementSitDAC;

end;

procedure TOF_BTREGLPIECE_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,
  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,
  Tiers_: THEdit);
begin
  inherited;
  Aff:=THEdit(GetControl('GP_AFFAIRE'))   ;
  Aff0:=THEdit(GetControl('GP_AFFAIRE0'));
  Aff1:=THEdit(GetControl('GP_AFFAIRE1')) ;
  Aff2:=THEdit(GetControl('GP_AFFAIRE2')) ;
  Aff3:=THEdit(GetControl('GP_AFFAIRE3')) ;
  Aff4:=THEdit(GetControl('GP_AVENANT'))  ;
  Tiers:=THEdit(GetControl('GP_TIERS'))   ;
end;


procedure TOF_BTREGLPIECE_MUL.DemandeReglementSitDAC; //( parms: array of variant; nb: integer ) ;
var Tiers       : String;
    Cledoc      : R_Cledoc;
begin

{$IFDEF EAGLCLIENT}
    TFMul(ecran).Q.TQ.Seek(TFMul(ecran).FListe.Row-1) ;
    Cledoc.naturePiece  := TFMul(ecran).Q.FindField('GP_NATUREPIECEG').AsString;
    Cledoc.DatePiece    := TFMul(ecran).Q.FindField('GP_DATEPIECE').AsDateTime;
    Cledoc.souche       := TFMul(ecran).Q.FindField('GP_SOUCHE').AsString;
    Cledoc.numeroPiece  := StrToInt(TFMul(ecran).Q.FindField('GP_NUMERO').AsString);
    Cledoc.indice       := StrToInt(TFMul(ecran).Q.FindField('GP_INDICEG').AsString);
{$ELSE}
    Cledoc.naturePiece  := Fliste.datasource.dataset.FindField('GP_NATUREPIECEG').AsString;
    Cledoc.DatePiece    := Fliste.datasource.dataset.FindField('GP_DATEPIECE').AsDateTime;
    Cledoc.souche       := Fliste.datasource.dataset.FindField('GP_SOUCHE').AsString;
    Cledoc.numeroPiece  := StrToInt(Fliste.datasource.dataset.FindField('GP_NUMERO').AsString);
    Cledoc.indice       := StrToInt(Fliste.datasource.dataset.FindField('GP_INDICEG').AsString);
    Tiers               := Fliste.datasource.dataset.FindField('GP_TIERS').AsString;
{$ENDIF}

  SaisieReglementsPiece(Cledoc, Tiers);

end;


procedure TOF_BTREGLPIECE_MUL.SaisieReglementsPiece(Cledoc : R_Cledoc; Tiers : String);
var XX : TAcceptationDocument;
		Xp,XD : double;
begin

  XX := TAcceptationDocument.create;

  XX.Fgestion := TmgIntervenant;

  if TypeReglt= 'COTRAITANCE' then
    XX.Fgestion := TmgCotraitance
  else if TypeReglt='SOUSTRAITANCE' then
    XX.Fgestion := TmgSousTraitance
  else if TypeReglt='INTERVENANT' then
    XX.Fgestion := TmgIntervenant;

  XX.ReglePiece := False;

  XX.naturePiece := CleDoc.NaturePiece;
  XX.Souche := CleDoc.Souche;
  XX.numero := CleDoc.NumeroPiece;
  XX.Indice := CleDoc.Indice;
  XX.Tiers := Tiers;
  //
  XX.ChargeLesTobs;
  GetSommeReglement (XX.TOBAcomptes ,Xp,XD,'',true);
  if (copy(TForm(ecran).Name,1,17) = 'BTREGLCOTRAIT_MUL') and (XD <> 0) then
  begin
  	PGIError('IMPOSSIBLE : Des règlements ont déjà été saisis sur ce document');
  end else
  begin
    //
    if TForm(ecran).Name = 'BTREGLPIECE_MUL' then
      XX.ReglePiece := True
    else if TForm(ecran).Name = 'BTREGLCOTRAIT_MUL' then
      XX.ReglePiece := False;
    //
    XX.GereReglements;
  end;
  XX.free;
end;

procedure TOF_BTREGLPIECE_MUL.ControleChamp(Champ, Valeur: String);
begin

  TypeReglt := '';

  if champ = 'COTRAITANCE' then
  begin
    TypeReglt := Champ;
    Ecran.Caption := 'Gestion des Réglements Cotraitants';
    If Assigned(GetControl('XX_WHERE')) then Setcontroltext('XX_WHERE', 'AND AFF_MANDATAIRE IN ("X", "-")');
  end
  else if champ = 'SOUSTRAITANCE' then
  begin
    TypeReglt := Champ;
    Ecran.Caption := 'Gestion des Réglements sous-Traitants';
    If Assigned(GetControl('XX_WHERE')) then Setcontroltext('XX_WHERE', 'AND SSTRAITE > 0');
  end
  else if champ = 'INTERVENANT' then
  begin
    TypeReglt := Champ;
    Ecran.Caption := 'Gestion des Réglements Intervenants';
    If Assigned(GetControl('XX_WHERE')) then Setcontroltext('XX_WHERE', 'AND AFF_MANDATAIRE IN ("X", "-") OR SSTRAITE > 0');
  end;

  //
  If Champ = 'STATUT' then
  Begin
    if Valeur = 'APP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'W')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'W');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="W"');
      SetControlproperty ('GP_NATUREPIECEG','Plus','AND GPP_NATUREPIECEG="FAC"');
      If Assigned(GetControl('XX_WHERE1')) then
      begin
        ThEdit(GetControl('XX_WHERE1')).Text := ThEdit(GetControl('XX_WHERE1')).Text +
                                              ' AND GP_NATUREPIECEG IN ("FAC")';
      end;
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
      SetControlText('TGP_AFFAIRE1', 'Appel');
    end
    Else if valeur = 'INT' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'I')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'I');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1)="I"');
      SetControlproperty ('GP_NATUREPIECEG','Plus','AND GPP_NATUREPIECEG="FAC"');
      If Assigned(GetControl('XX_WHERE1')) then
      begin
        ThEdit(GetControl('XX_WHERE1')).Text := ThEdit(GetControl('XX_WHERE1')).Text +
                                              ' AND GP_NATUREPIECEG IN ("FAC")';
      end;
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
      SetControlText('TGP_AFFAIRE1', 'Contrat');
    end
    Else if valeur = 'AFF' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'A')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'A');
      SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A", "")');
      SetControlproperty ('GP_NATUREPIECEG','Plus','AND GPP_NATUREPIECEG="FBT" OR GPP_NATUREPIECEG="DAC" OR GPP_NATUREPIECEG="FAC"');
      If Assigned(GetControl('XX_WHERE1')) then
      begin
        ThEdit(GetControl('XX_WHERE1')).Text := ThEdit(GetControl('XX_WHERE1')).Text +
                                              ' AND GP_NATUREPIECEG IN ("FBT","DAC","FAC")';
      end;
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
      SetControlText('TGP_AFFAIRE1', 'Chantier');
    end
    else if Valeur = 'GRP' then
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControltext('XX_WHERE',' AND AFF_AFFAIRE0 in ("A", "")')
      else if assigned(GetControl('AFFAIRE0')) then SetControltext('XX_WHERE',' AND SUBSTRING(GP_AFFAIRE,1,1) in ("A", "")');
    end
    Else if valeur = 'PRO' then
    Begin
      //
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', 'P')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', 'P');
      SetControltext('XX_WHERE',' AND SUBSTRING(AFF_AFFAIRE,1,1)="P"');
      //
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
      SetControlText('TGP_AFFAIRE1', 'Appel d''offre');
    end
    else
    Begin
      if assigned(GetControl('AFF_AFFAIRE0'))  then SetControlText('AFF_AFFAIRE0', '')
      else if assigned(GetControl('AFFAIRE0')) then SetControlText('AFFAIRE0', '');
      SetControltext('XX_WHERE','');
      SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
      SetControlText('TGP_AFFAIRE1', 'Affaire');
    end
  end;

end;

Initialization
  registerclasses ( [ TOF_BTREGLPIECE_MUL ] ) ;
 	//RegisterAglProc( 'DemandeReglementSitDAC',True,6,DemandeReglementSitDAC);
end.

