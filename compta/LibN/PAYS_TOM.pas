{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 27/02/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PAYS (PAYS)
Mots clefs ... : TOM;PAYS
*****************************************************************}
Unit PAYS_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     eFiche,        // TTFiche
     MaineAGL,      // AGLLanceFiche
     UtileAGL,      // TNavigateBtn
     eTablette,
{$ELSE}
     db,
     dbtables,
     HDB,
     FE_Main,  // AGLLanceFiche
     dbCtrls,  // Contrôles DB
     Fiche, 
     FichList,
{$ENDIF}
     utilPGI,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,Dialogs,
     UTob,
     CPREGION_TOF ;

Type
  TOM_PAYS = Class (TOM)
    PY_PAYS: THEdit;
    PY_ABREGE: THEdit;
    PY_LIBELLE: THEdit;
    PY_REGION: TCheckBox;
    PY_CODEISO2: THEdit;
    TPY_CODEBANCAIRE: THLabel;
    PY_CODEBANCAIRE: THEdit;
    PY_MEMBRECEE : THEdit ;
    PY_DEVISE : THEdit ;
    PY_DRAPEAUX : THEdit ;
    PY_NATIONALITE : THEdit ;
    PY_LANGUE : THEdit ;
    procedure PY_REGIONOnClick(Sender: TObject);
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  private    { Déclarations privées }
    MemoRegion : Boolean ;
    Acreer : Boolean ;
    QuelPays : String ;
  public    { Déclarations publiques }
end ;

  Procedure OuvrePAYS ;


Implementation

Procedure OuvrePAYS ;
begin
  if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
  AGLLanceFiche('YY','YYPAYS','','','ACTION=MODIFICATION');
END ;

procedure TOM_PAYS.PY_REGIONOnClick(Sender: TObject);
begin
  inherited;
Acreer:=(PY_REGION.State=cbChecked)And(Not MemoRegion) ;
end;

procedure TOM_PAYS.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_PAYS.OnDeleteRecord ;
begin
  Inherited ;
  ExecuteSql('Delete from REGION Where RG_PAYS="'+GetField('PY_PAYS')+'"') ;
end ;

procedure TOM_PAYS.OnUpdateRecord ;
begin
  Inherited ;
if V_PGI.LaSerie>S2 then
   BEGIN
   if GetField('PY_CODEISO2')='' then
      BEGIN PY_CODEISO2.SetFocus ; Exit ; END ;
   if GetField('PY_CODEBANCAIRE')='' then
      BEGIN PY_CODEBANCAIRE.SetFocus ; Exit ; END ;
   END ;
  QuelPays:=GetField('PY_PAYS') ;
  if(Acreer) then //And(HM2.Execute(2,'','')=mrYes)then
    FicheRegion(QuelPays,'',False) ;
  //ParamTable('TTREGION',tacreat,nil,nil) ;
//    AGLLanceFiche( 'YY', 'CP','',quelPays,'ACTION=MODIFICATION');
  if(PY_REGION.State=cbUnChecked)then
    ExecuteSql('Delete from REGION Where RG_PAYS="'+GetField('PY_PAYS')+'"') ;

end ;

procedure TOM_PAYS.OnAfterUpdateRecord ;
begin
  Inherited ;


end ;

procedure TOM_PAYS.OnLoadRecord ;
begin
  Inherited ;
  MemoRegion:=(PY_REGION.State=cbChecked) ;
end ;

procedure TOM_PAYS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PAYS.OnArgument ( S: String ) ;
begin
  Inherited ;
  Ecran.HelpContext := 500009;

  PY_CODEISO2:= THEdit(GetControl('PY_CODEISO2'));
  PY_CODEBANCAIRE:= THEdit(GetControl('PY_CODEBANCAIRE'));
  TPY_CODEBANCAIRE:= THLabel(GetControl('TPY_CODEBANCAIRE'));
  PY_REGION := TCheckBox(GetControl('PY_REGION'));
  SetControlVisible('PY_CODEISO2',V_PGI.LaSerie>S2);
  SetControlvisible('TPY_CODEBANCAIRE',V_PGI.LaSerie>S2) ;
  SetControlVisible('PY_CODEBANCAIRE',V_PGI.LaSerie>S2) ;
  PY_REGION.OnClick:=PY_REGIONOnClick;
end ;

procedure TOM_PAYS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PAYS.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_PAYS ] ) ; 
end.
