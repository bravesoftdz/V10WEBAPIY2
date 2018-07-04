unit UtofRTPiece_TV;

interface
uses  StdCtrls,Classes,forms,sysutils,
      HEnt1,UTOF,
      UTobView, UtilSelection ,UtilRT ,EntGC, FactUtil,uEntCommun
;
Type
     TOF_RTPIECE_TV = Class (TOF)
     private
         TobViewer1: TTobViewer;
         procedure TVOnDblClickCell(Sender: TObject ) ;
     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
     END ;

implementation

uses
  UtilPGI,ParamSoc,Facture;

procedure TOF_RTPIECE_TV.OnArgument(Arguments : String ) ;
var F : TForm;
begin
inherited ;
  F := TForm (Ecran);
  MulCreerPagesCL(F,'NOMFIC=GCTIERS');
  if GetParamSocSecur('SO_RTGESTINFOS00V',False) then
    MulCreerPagesCL(Ecran,'NOMFIC=RTPERSPECTIVES');
  TobViewer1:=TTobViewer(getcontrol('TV'));
  TobViewer1.OnDblClick:= TVOnDblClickCell ;
end;

procedure TOF_RTPIECE_TV.OnLoad;
Var F : TForm;
    xx_where : string  ;
begin
inherited;
F := TForm (Ecran);
xx_where:=GetControlText('XX_WHERE');
if (TCheckbox(F.FindComponent('PROPOPRINCIPALE')).Checked = true) then
   xx_where := xx_where+' AND (RPE_VARIANTE=0 or RPE_PERSPECTIVE=RPE_VARIANTE)';
xx_where := xx_where + RTXXWhereConfident('CON');

SetControlText('XX_WHERE',xx_where) ;
end;

procedure TOF_RTPIECE_TV.TVOnDblClickCell(Sender: TObject );
var
  staction: string;
  CleDoc: R_CleDoc;
  Param: R_SaisiePieceParam;
begin
with TTobViewer(sender) do
  begin
  staction:='ACTION=CONSULTATION';
  if RTDroitModifTiers(AsString[ColIndex('GP_TIERS'), CurrentRow]) then staction:='ACTION=MODIFICATION';
  if ((ColName[CurrentCol] = 'RPE_AUXILIAIRE') or (ColName[CurrentCol] = 'RPE_TIERS')
    or (ColName[CurrentCol] = 'GP_TIERS') or (ColName[CurrentCol] = 'T_TIERS')
    or (ColName[CurrentCol] = 'T_LIBELLE') or (ColName[CurrentCol] = 'RaisonSociale')) then
       V_PGI.DispatchTT (28,taConsult ,AsString[ColIndex('RPE_AUXILIAIRE'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'RPE_NUMEROACTION') and  (AsInteger[ColIndex('RPE_NUMEROACTION'), CurrentRow] <> 0 )
  then V_PGI.DispatchTT (22,taConsult ,AsString[ColIndex('RPE_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('RPE_NUMEROACTION'), CurrentRow]), '','')
  else if (ColName[CurrentCol] = 'RPE_OPERATION') and (trim(AsString[ColIndex('RPE_OPERATION'), CurrentRow])<>'')
  then  V_PGI.DispatchTT (23,taConsult ,AsString[ColIndex('RPE_OPERATION'), CurrentRow], '','')

  else if ((ColName[CurrentCol] = 'RPE_PROJET') and (trim(AsString[ColIndex('RPE_PROJET'), CurrentRow])<>'')) or
       ((ColName[CurrentCol] = 'RPJ_PROJET') and (trim(AsString[ColIndex('RPJ_PROJET'), CurrentRow])<>'')) or
       (ColName[CurrentCol] = 'RPJ_LIBELLE')
  then V_PGI.DispatchTT (30,taConsult ,AsString[ColIndex('RPE_PROJET'), CurrentRow], '','')

  else if ((ColName[CurrentCol] = 'RPE_PERSPECTIVE') and (AsInteger[ColIndex('RPE_PERSPECTIVE'), CurrentRow]<>0)) or
          (ColName[CurrentCol] = 'GP_PERSPECTIVE') and (AsInteger[ColIndex('GP_PERSPECTIVE'), CurrentRow]<>0) or
          (ColName[CurrentCol] = 'RPE_LIBELLE')
  then V_PGI.DispatchTT (40,taConsult ,IntToStr(AsInteger[ColIndex('GP_PERSPECTIVE'), CurrentRow]), '','')

  else if (ColName[CurrentCol] = 'T_SOCIETEGROUPE')  then  V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex('T_SOCIETEGROUPE'), CurrentRow], '','')
  else if (ColName[CurrentCol] = 'T_PRESCRIPTEUR')  then  V_PGI.DispatchTT (8,taConsult ,AsString[ColIndex('T_PRESCRIPTEUR'), CurrentRow], '','')
  else
    begin
//'FIXED;GP_NATUREPIECEG;GP_NUMERO;GP_DATEPIECE;GP_TOTALHT;GP_DEVISE;GP_SOUCHE;GP_INDICEG;GP_NATUREPIECEG';
    DecodeRefPiece (
        AsString[ColIndex('GP_NATUREPIECEG'), CurrentRow] + ';' + AsString[ColIndex('GP_SOUCHE'), CurrentRow] + ';' +
             IntToStr(AsInteger[ColIndex('GP_NUMERO'), CurrentRow]) +
        ';' + IntToStr(AsInteger[ColIndex('GP_INDICEG'), CurrentRow]) + ';',
        CleDoc);
    CleDoc.DatePiece:=AsDateTime[ColIndex('GP_DATEPIECE'), CurrentRow];
//    InitRecordSaisiePieceParam(Param);
    SaisiePiece (CleDoc, taConsult);
    end;
  end;
end;

Initialization
registerclasses([TOF_RTPIECE_TV]);

end.
