{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/05/2016
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSELDEVIS_MUL ()
Mots clefs ... : TOF;BTSELDEVIS_MUL
*****************************************************************}
Unit BTSELDEVIS_MUL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     uTob,
     HTB97,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UTofAfBaseCodeAffaire,
     UentCommun,
     UtilTOBpiece,
     AffaireUtil
      ;

Type
  TOF_BTSELDEVIS_MUL =  Class (TOF_AFBASECODEAFFAIRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2, Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers, Tiers_: THEdit); override;
  private
    procedure FlisteDblClick (Sender : TOBject);
    procedure ControleChamp(Champ, Valeur: String);
  end ;

Implementation

procedure TOF_BTSELDEVIS_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSELDEVIS_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSELDEVIS_MUL.OnUpdate ;
begin
  Inherited ;
  if ThEdit(getControl('GP_TIERSFACTURE')).Text<> '' then ThEdit(getControl('GP_TIERSFACTURE')).Enabled := false;
  if ThEdit(getControl('GP_AFFAIRE')).Text<> '' then
  begin
    ChargeCleAffaire(ThEdit(GetControl('AFFAIRE0')), ThEdit(GetControl('GP_AFFAIRE1')), ThEdit(GetControl('GP_AFFAIRE2')), ThEdit(GetControl('GP_AFFAIRE3')),
                           ThEdit(GetControl('GP_AVENANT')), nil, taConsult, ThEdit(GetControl('GP_AFFAIRE')).Text, False);
    ThEdit(getControl('AFFAIRE0')).Enabled := false;
    ThEdit(getControl('GP_AFFAIRE1')).Enabled := false;
    ThEdit(getControl('GP_AFFAIRE2')).Enabled := false;
    ThEdit(getControl('GP_AFFAIRE3')).Enabled := false;
    ThEdit(getControl('GP_AVENANT')).Enabled := false;
    TToolbarButton97(getControl('BSELECTAFF1')).Enabled := false;
    TToolbarButton97(getControl('BEFFACEAFF1')).Enabled := false;
  end;
end ;

procedure TOF_BTSELDEVIS_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTSELDEVIS_MUL.OnArgument (S : String ) ;
var critere,StArgument : string;
    X : integer;
    valmul,ChampMul : string;
begin
  fMulDeTraitement  := true;
  Inherited ;

  fTableName        := 'PIECE';
  SetControlText('XX_WHERE', '');
  StArgument := S;
  Repeat
    Critere:=uppercase(ReadTokenSt(stArgument)) ;
    valMul := '';
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        ChampMul:=copy(Critere,1,x-1);
        ValMul:=copy(Critere,x+1,length(Critere));
      end
      else
        ChampMul := Critere;
      ControleChamp(ChampMul, ValMul);
    end;
  until  Critere='';


  THGrid(GetControl('Fliste')).OnDblClick := FlisteDblClick;

end ;

procedure TOF_BTSELDEVIS_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTSELDEVIS_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELDEVIS_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSELDEVIS_MUL.NomsChampsAffaire(var Aff, Aff0, Aff1, Aff2,  Aff3, Aff4, Aff_, Aff0_, Aff1_, Aff2_, Aff3_, Aff4_, Tiers,  Tiers_: THEdit);
begin
  inherited;
  Aff0 := THEdit(GetControl('AFFAIRE0'));
  Aff := THEdit(GetControl('GP_AFFAIRE'));
  Aff1 := THEdit(GetControl('GP_AFFAIRE1'));
  Aff2 := THEdit(GetControl('GP_AFFAIRE2'));
  Aff3 := THEdit(GetControl('GP_AFFAIRE3'));
  Aff4 := THEdit(GetControl('GP_AVENANT'));
end;

procedure TOF_BTSELDEVIS_MUL.FlisteDblClick(Sender: TOBject);
var NaturepieceG,Souche : String;
    indiceG,Numero : integer;
    Zcledoc : r_cledoc;
begin

  With TFMul(Ecran) do
  begin
    {$IFDEF EAGLCLIENT}
    Q.TQ.Seek(FListe.Row-1) ;
    {$ENDIF}
    NaturepieceG := Q.FindField('GP_NATUREPIECEG').asstring ;
    Souche := Q.FindField('GP_SOUCHE').asstring ;
    Numero := Q.FindField('GP_NUMERO').asInteger ;
    IndiceG := Q.FindField('GP_INDICEG').asInteger ;
    Zcledoc.NaturePiece := NaturepieceG;
    Zcledoc.Souche := Souche;
    Zcledoc.NumeroPiece := Numero;
    Zcledoc.Indice := IndiceG;
    if Naturepieceg='DBT' then
    begin
      if ExisteSql('SELECT 1 FROM LIGNE WHERE '+WherePiece(ZCleDoc, ttdLigne, False)+' AND ((GL_QTEPREVAVANC) <> 0 or (GL_QTESIT <> 0))') then
      begin
        PgiError('Impossible ce devis a déjà été facturé');
        exit;
      end;
    end;
    LaTOB.SetString('RESULT', '01011900;' + NaturepieceG  + ';' + Souche  + ';' +IntToStr(Numero) + ';' + IntToStr(IndiceG) + ';;');
    ecran.close;
  end;

end;

Procedure TOF_BTSELDEVIS_MUL.ControleChamp(Champ : String;Valeur : String);
Begin

  if champ='ETAT' then SetControlText('AFF_ETATAFFAIRE', Valeur);

  if Champ='STATUT' then
     Begin
     if Valeur = 'APP' then
        Begin
	    	SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("","W")');
        SetControlText('AFFAIRE0', 'W');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel');
        SetControlText('TGPAFFAIRE', 'Code Appel');
        end
     else if Valeur = 'INT' then
        Begin
    		SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND ((SUBSTRING(GP_AFFAIRE,1,1)="I") OR (GP_GENERAUTO="CON"))');
        SetControlText('AFFAIRE0', 'I');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Contrat');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Contrat');
        SetControlText('TGPAFFAIRE', 'Code Contrat');
        End
     else if Valeur = 'GRP' then
        Begin
	    	SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1)IN ("W","A")');
        SetControlText('AFFAIRE0', 'A');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
        SetControlText('TGPAFFAIRE', 'Code Affaire');
        end
     else if Valeur = 'AFF' then
        Begin
      	SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1) IN ("A","") AND GP_ATTACHEMENT=""');
        SetControlText('AFFAIRE0', 'A');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Chantier');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Chantier');
        SetControlText('TGPAFFAIRE', 'Code Chantier');
        end
     else if Valeur = 'PRO' then
        Begin
      	SetControlText('XX_WHERE', GetControltext('XX_WHERE')+' AND SUBSTRING(GP_AFFAIRE,1,1)="P"');
        SetControlText('AFFAIRE0', 'P');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Appel d''offre');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Appel d''offre');
        SetControlText('TGPAFFAIRE', 'Code Appel d''offre');
        end
     else
        Begin
        SetControlText('AFFAIRE0', '');
        SetControlProperty('BSELECTAFF1', 'Hint', 'Recherche Affaire');
        SetControlProperty('BEFFACEAFF1', 'Hint', 'Effacer Critères Affaire');
        SetControlText('TGPAFFAIRE', 'Code Affaire');
        end;
     end;

  if Champ='AFF_GENERAUTO' then
  begin
    SetControltext('AFF_GENERAUTO',Valeur);
  end;

end;


Initialization
  registerclasses ( [ TOF_BTSELDEVIS_MUL ] ) ;
end.
