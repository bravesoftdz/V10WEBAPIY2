{***********UNITE*************************************************
Auteur  ...... : ROHAULT Régis
Créé le ...... : 08/09/2006
Modifié le ... : 12/09/2006
Description .. : Source TOM de la TABLE : YIDENTBANCAIRE
Suite ........ : (YIDENTBANCAIRE)
Suite ........ :
Suite ........ : permet la gestion de la fiche d'identification bancaire.
Suite ........ : paramétrage des longueurs des zones et leurs ordres afin de
Suite ........ : constituer un numéro de compte complet et à la norme d'un
Suite ........ : pays
Mots clefs ... : TOM;YIDENTBANCAIRE
*****************************************************************}
unit TomIdentificationBancaire;

interface

uses StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
  db,DBCtrls,
  {$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  Fiche,
  FichList,
  FE_Main, // AGLLanceFiche
  {$ELSE}
  eFiche,
  eFichList,
  MaineAGL, // AGLLanceFiche
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOM,
  HTB97,
  wCommuns,
  ULibIdentBancaire,
  utobdebug,
  UtilPGI,
  UTob,
  HDB;

procedure LanceFicheIdentificationBancaire;

type
  TOM_YIDENTBANCAIRE = class(TOM)

    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnAfterDeleteRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;

  private
  {$IFNDEF EAGLCLIENT}
    YIB_NETABBQ: THDBSpinEdit;
    YIB_NGUICHET: THDBSpinEdit;
    YIB_NNUMEROCOMPTE: THDBSpinEdit;
    YIB_NCLERIB: THDBSpinEdit;
    YIB_LGCODEIBAN : THDBSpinEdit;
    YIB_LGETABBQ: THDBSpinEdit;
    YIB_LGGUICHET: THDBSpinEdit;
    YIB_LGNUMEROCOMPTE: THDBSpinEdit;
    YIB_LGCLERIB: THDBSpinEdit;
    YIB_PAYSISO: THDBEdit;
    YIB_TYPEPAYS : THDBEdit;
    YIB_TYPECOMPTE : THDBValComboBox  ;
  {$ELSE}
    YIB_NETABBQ: THSpinEdit;
    YIB_NGUICHET: THSpinEdit;
    YIB_NNUMEROCOMPTE: THSpinEdit;
    YIB_NCLERIB: THSpinEdit;
    YIB_LGCODEIBAN : THSpinEdit;
    YIB_LGETABBQ: THSpinEdit;
    YIB_LGGUICHET: THSpinEdit;
    YIB_LGNUMEROCOMPTE: THSpinEdit;
    YIB_LGCLERIB: THSpinEdit;
    YIB_PAYSISO: THEdit;
    YIB_TYPEPAYS : THEdit;
    YIB_TYPECOMPTE : THValComboBox  ;
  {$ENDIF}
    BGParam : TGroupBox ;
    procedure lanceAffichage ;
    procedure OnChangeTypeCompte(Sender: TObject);
  end;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /    
Description .. : Affiche la fiche YYIDENTBANCAIRE
Mots clefs ... : 
*****************************************************************}
procedure LanceFicheIdentificationBancaire;
begin
  AGLLanceFiche('YY', 'YYIDENTBANCAIRE', '', '', '');
end;

procedure TOM_YIDENTBANCAIRE.OnNewRecord;
begin
  inherited;
end;

procedure TOM_YIDENTBANCAIRE.OnDeleteRecord;
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Regis Rohault
Créé le ...... : 20/11/2006
Modifié le ... :   /  /    
Description .. : Verifie avant la mise a jour la concordance des données 
Suite ........ : des composants du RIB
Mots clefs ... : 
*****************************************************************}
procedure TOM_YIDENTBANCAIRE.OnUpdateRecord;
begin
  inherited;
  // Vérification des zones
  //SDA le 14/05/2007
  if trim(YIB_TYPEPAYS.Text) = '' then
  begin
    MESSAGEALERTE('Merci de renseigner le type d''identification.');
    SetFocusControl('YIB_TYPEPAYS');
    LastError := 1;
  end;
  //Fin SDA le 14/05/2007
  
  if YIB_TYPECOMPTE.value='RIB' then
  begin
    if YIB_NETABBQ.value = 0 then
    begin
      MESSAGEALERTE('La position du code banque doit être comprise entre 1 et 4');
      SetFocusControl('YIB_NETABBQ');
      LastError := 1;
    end
    else if YIB_NGUICHET.value = 0 then
    begin
        MESSAGEALERTE('La position du code guichet doit être compris entre 1 et 4');
        SetFocusControl('YIB_NGUICHET');
        LastError := 1;
    end
    else if YIB_NNUMEROCOMPTE.value = 0 then
    begin
        MESSAGEALERTE('La position du numéro de compte doit être compris entre 1 et 4');
        SetFocusControl('YIB_NNUMEROCOMPTE');
        LastError := 1;
    end
    else if YIB_NCLERIB.value = 0 then
    begin
        MESSAGEALERTE('La position de la clé RIB doit être compris entre 1 et 4');
        SetFocusControl('YIB_NCLERIB');
        LastError := 1;
    end
    else if YIB_NCLERIB.value + YIB_NNUMEROCOMPTE.value + YIB_NETABBQ.value + YIB_NGUICHET.value <> 10 then
    begin
      MESSAGEALERTE('ATTENTION ! Les numéros d''ordre sont erronés !');
      SetFocusControl('YIB_NETABBQ');
      LastError := 1;
    end
    else if YIB_LGCLERIB.value + YIB_LGNUMEROCOMPTE.value + YIB_LGETABBQ.value + YIB_LGGUICHET.value < 5 then
    begin
      MESSAGEALERTE('ATTENTION ! Les longueurs des zones sont erronés !');
      SetFocusControl('YIB_LGETABBQ');
      LastError := 1;
    end;
  end ;
  //lanceAffichage;
end;


procedure TOM_YIDENTBANCAIRE.lanceAffichage;
begin
  If (YIB_TYPECOMPTE.Value='BBA') or (YIB_TYPECOMPTE.Value='IBA') then
  begin
    SetControlVisible('BGParam',False) ;
    SetControlVisible('YIB_LGCODEIBAN',True) ;
    (*
    SetControlVisible('TYIB_LGETABBQ',False) ;
    SetControlVisible('TYIB_LGGUICHET',False) ;
    SetControlVisible('TYIB_LGCOMPTE',False) ;
    SetControlVisible('TYIB_LGCLE',False) ;

    SetControlVisible('YIB_LGETABBQ',False) ;
    SetControlVisible('YIB_LGGUICHET',False) ;
    SetControlVisible('YIB_LGNUMEROCOMPTE',False) ;
    SetControlVisible('YIB_LGCLERIB',False) ;

    SetControlVisible('TYIB_NUMETABBQ',False) ;
    SetControlVisible('TYIB_NUMGUICHET',False) ;
    SetControlVisible('TYIB_NUMCOMPTE',False) ;
    SetControlVisible('TYIB_NUMCLERIB',False) ;

    SetControlVisible('YIB_NETABBQ',False) ;
    SetControlVisible('YIB_NGUICHET',False) ;
    SetControlVisible('YIB_NNUMEROCOMPTE',False) ;
    SetControlVisible('YIB_NCLERIB',False) ;
     *)
  end
  else
  begin
  (*
    SetControlVisible('TYIB_LGETABBQ',True) ;
    SetControlVisible('TYIB_LGGUICHET',True) ;
    SetControlVisible('TYIB_LGCOMPTE',True) ;
    SetControlVisible('TYIB_LGCLE',True) ;

    SetControlVisible('YIB_LGETABBQ',True) ;
    SetControlVisible('YIB_LGGUICHET',True) ;
    SetControlVisible('YIB_LGNUMEROCOMPTE',True) ;
    SetControlVisible('YIB_LGCLERIB',True) ;

    SetControlVisible('TYIB_NUMETABBQ',True) ;
    SetControlVisible('TYIB_NUMGUICHET',True) ;
    SetControlVisible('TYIB_NUMCOMPTE',True) ;
    SetControlVisible('TYIB_NUMCLERIB',True) ;

    SetControlVisible('YIB_NETABBQ',True) ;
    SetControlVisible('YIB_NGUICHET',True) ;
    SetControlVisible('YIB_NNUMEROCOMPTE',True) ;
    SetControlVisible('YIB_NCLERIB',True) ;
  *)
    //SDA le 22/05/2007 problème affichage
    SetControlVisible('BGParam',True) ;
    SetControlVisible('YIB_LGCODEIBAN',false);
    //Fin SDA le 22/05/2007

    //YIB_LGCODEIBAN.Value := 0 ;
  end ;
end;

procedure TOM_YIDENTBANCAIRE.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_YIDENTBANCAIRE.OnAfterDeleteRecord;
begin
  inherited;
end;

procedure TOM_YIDENTBANCAIRE.OnLoadRecord;
begin
  inherited;
  if DS.state in [dsEdit, dsBrowse]  then
    lanceAffichage;
end;

procedure TOM_YIDENTBANCAIRE.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_YIDENTBANCAIRE.OnArgument(S: string);
begin
  inherited;
  {$IFNDEF EAGLCLIENT} //SDA le 22/05/2007 prise en compte EAGLCLIENT
  YIB_TYPEPAYS := THDBEdit(GetControl('YIB_TYPEPAYS', true)); //SDA le 14/05/2007
  YIB_PAYSISO := THDBEdit(GetControl('YIB_PAYSISO', true));
  YIB_NETABBQ := THDBSpinEdit(GetControl('YIB_NETABBQ', true));
  YIB_NGUICHET := THDBSpinEdit(GetControl('YIB_NGUICHET', true));
  YIB_NNUMEROCOMPTE := THDBSpinEdit(GetControl('YIB_NNUMEROCOMPTE', true));
  YIB_NCLERIB := THDBSpinEdit(GetControl('YIB_NCLERIB', true));
  YIB_LGCODEIBAN := THDBSpinEdit(GetControl('YIB_LGCODEIBAN', true));
  YIB_LGETABBQ := THDBSpinEdit(GetControl('YIB_LGETABBQ', true));
  YIB_LGGUICHET := THDBSpinEdit(GetControl('YIB_LGGUICHET', true));
  YIB_LGNUMEROCOMPTE := THDBSpinEdit(GetControl('YIB_LGNUMEROCOMPTE', true));
  YIB_LGCLERIB := THDBSpinEdit(GetControl('YIB_LGCLERIB', true));
  YIB_TYPECOMPTE := THDBValComboBox(GetControl('YIB_TYPECOMPTE', true));
  {$ELSE}
  YIB_TYPEPAYS := THEdit(GetControl('YIB_TYPEPAYS', true)); //SDA le 14/05/2007
  YIB_PAYSISO := THEdit(GetControl('YIB_PAYSISO', true));
  YIB_NETABBQ := THSpinEdit(GetControl('YIB_NETABBQ', true));
  YIB_NGUICHET := THSpinEdit(GetControl('YIB_NGUICHET', true));
  YIB_NNUMEROCOMPTE := THSpinEdit(GetControl('YIB_NNUMEROCOMPTE', true));
  YIB_NCLERIB := THSpinEdit(GetControl('YIB_NCLERIB', true));
  YIB_LGCODEIBAN := THSpinEdit(GetControl('YIB_LGCODEIBAN', true));
  YIB_LGETABBQ := THSpinEdit(GetControl('YIB_LGETABBQ', true));
  YIB_LGGUICHET := THSpinEdit(GetControl('YIB_LGGUICHET', true));
  YIB_LGNUMEROCOMPTE := THSpinEdit(GetControl('YIB_LGNUMEROCOMPTE', true));
  YIB_LGCLERIB := THSpinEdit(GetControl('YIB_LGCLERIB', true));
  YIB_TYPECOMPTE := THValComboBox(GetControl('YIB_TYPECOMPTE', true));
  {$ENDIF}

  YIB_TYPECOMPTE.OnChange := OnChangeTypeCompte ;
  BGParam := TGroupBox(GetControl('BGParam',true));

end;

procedure TOM_YIDENTBANCAIRE.OnClose;
begin
  inherited;
end;

procedure TOM_YIDENTBANCAIRE.OnCancelRecord;
begin
  inherited;
end;

procedure TOM_YIDENTBANCAIRE.OnChangeTypeCompte(Sender: TObject);
begin
  lanceAffichage;
end;

initialization
  registerclasses([TOM_YIDENTBANCAIRE]);
end.

