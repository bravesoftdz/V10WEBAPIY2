{***********UNITE*************************************************
Auteur  ...... : Agnès Cathelineau
Créé le ...... : 30/07/2001
Modifié le ... :   /  /
Description .. : Spécifique mode: Tarif détail
Suite ........ : Gestion des périodes d'application pour les tarifs
Mots clefs ... : TARIF;PERIODE
*****************************************************************}
unit UtomTarifPer;

interface

uses  UTOM,HCtrls,Classes,Controls,Sysutils,HDB,Fiche,DBTables,HEnt1,M3FP,DB,Forms,LookUp,
      HMsgBox,TarifUtil,Fe_Main,FichList,UTob ;

Type
     TOM_TarifPer = Class (TOM)

     Private
       procedure SetLastError (Num : integer; ou : string );

     public
       Action   : TActionFiche ;
       procedure OnLoadRecord ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnUpdateRecord  ; override ;
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnDeleteRecord ; override ;
       function  PeriodeUtiliseDansTarif: Boolean ;
       Procedure ModifDateEtab ;
       Procedure MiseAJourTarif ;


     END ;

procedure SupprimePerEtab(CodePeriode : string) ;
Function RechTarfMode(periode:String ): String ;

const TexteMessage: array[1..6] of string 	= (
          {1}        'Vous devez renseigner une période d''application',
          {2}        'Vous devez renseigner un libellé',
          {3}        'Les bornes de dates de validité sont incohérentes',
          {4}        'Le code de la période d''application existe déjà',
          {5}        'Cette période est utilisé dans un tarif, vous ne pouvez pas la supprimer',
          {6}        'Il existe des remises sur cette période,impossible de la modifier en période de base'
          );

implementation

// Récupération de la nature de pièce en argument
procedure TOM_TarifPer.OnArgument(stArgument : String );
var x : integer;
    critere: string;
    Arg,Val : string;
BEGIN
Inherited;
Critere := (Trim(ReadTokenSt(stArgument)));
if Critere<>'' then
    BEGIN
    x:= pos('=', Critere);
    if x<>0 then
       BEGIN
       Arg:= copy (Critere, 1, x-1);
       Val:= copy (Critere, x+1, length(Critere));
       if Arg='ACTION' then
          if Val='CREATION' then Action:=taCreat ;
       END;
    END;
END;

procedure TOM_TarifPer.OnLoadRecord ;
var F : TFFicheListe ;
begin
Inherited ;
if GetField('GFP_PROMO')='X' then
  begin
  SetControlEnabled('GFP_CASCADE',True) ;
  SetControlEnabled('GFP_DATEFIN',True) ;
  SetControlEnabled('GFP_ARRONDI',True) ;
  SetControlEnabled('GFP_DEMARQUE',True) ;
  end else
  begin
  SetControlEnabled('GFP_CASCADE',False) ;
  SetControlEnabled('GFP_DATEFIN',False) ;
  SetControlEnabled('GFP_ARRONDI',False) ;
  SetControlEnabled('GFP_DEMARQUE',False) ;
  end ;
SetControlEnabled('GFP_CODEPERIODE',(DS.State in [dsInsert])) ;
SetControlEnabled('BETABLIS',(not (DS.State in [dsInsert]) and (GetField('GFP_PROMO')='X'))) ;
if DS.State in [dsInsert] then SetFocusControl('GFP_CODEPERIODE') else
  begin
  F := TFFicheListe(TForm(Ecran)) ;
  F.FRange :='...' ;
  SetLeRange(F.Ta,F.FRange);
  end ;
end ;

procedure TOM_TarifPer.OnUpdateRecord;
var SQLEtab: String ;
BEGIN
Inherited;
If GetField('GFP_CODEPERIODE')='' then
   begin SetLastError(1, 'GFP_CODEPERIODE'); exit ; end
   else if (DS.State in [dsInsert]) and (ExisteSQL('Select * from TARIFPER where GFP_CODEPERIODE="'+GetField('GFP_CODEPERIODE')+'" AND GFP_ETABLISSEMENT=""')) then
      begin SetLastError(4, 'GFP_CODEPERIODE'); exit ; end ;

If GetField('GFP_LIBELLE')='' then
   begin SetLastError(2, 'GFP_LIBELLE'); exit ; end ;

If GetField('GFP_DATEDEBUT') > GetField('GFP_DATEFIN') then
    begin SetLastError(3, 'GFP_DATEFIN'); exit ; end ;

// MAJ Table tarif mode
if ExisteSQL('SELECT GFM_TARFMODE FROM TARIFMODE WHERE GFM_PERTARIF="'+GetField('GFP_CODEPERIODE')+'"')
   then ExecuteSQL('Update TarifMode SET GFM_DATEDEBUT="'+USDateTime(GetField('GFP_DATEDEBUT'))+'",'+
           'GFM_DATEFIN="'+USDateTime(GetField('GFP_DATEFIN'))+'",'+
           'GFM_PROMO="'+GetField('GFP_PROMO')+'",'+
           'GFM_ARRONDI="'+GetField('GFP_ARRONDI')+'",'+
           'GFM_CASCADE="'+GetField('GFP_CASCADE')+'",'+
           'GFM_DEMARQUE="'+GetField('GFP_DEMARQUE')+'" where GFM_PERTARIF="'+GetField('GFP_CODEPERIODE')+'"') ;

// MAJ des periodes par établissement
SQLEtab:='Select GFP_ETABLISSEMENT from TARIFPER Where GFP_ETABLISSEMENT<>"..." And GFP_CODEPERIODE="'+GetField('GFP_CODEPERIODE')+'"' ;
if ExisteSQL(SQLEtab) then
  begin
  ExecuteSQL('Update TARIFPER SET GFP_PROMO="'+GetField('GFP_PROMO')+'",'+
           'GFP_CASCADE="'+GetField('GFP_CASCADE')+'",'+
           'GFP_DEMARQUE="'+GetField('GFP_DEMARQUE')+'",'+
           'GFP_LIBELLE="'+GetField('GFP_LIBELLE')+'" where GFP_CODEPERIODE="'+GetField('GFP_CODEPERIODE')+'" AND GFP_ETABLISSEMENT<>"..."') ;
  end ;
MiseAJourTarif ;
END;

procedure TOM_TarifPer.OnChangeField(F : TField) ;
var TarfMode:String ;
Begin
Inherited;
if F.FieldName='GFP_PROMO' then
   if GetField('GFP_PROMO')<>'X' then
   begin
     TarfMode:=RechTarfMode(GetControlText('GFP_CODEPERIODE')) ;
     If (TarfMode<>'') and (ExisteSQL('Select GF_REMISE From tarif where GF_TARFMODE IN ('+TarfMode+') AND GF_REMISE<>"0" AND GF_ARTICLE<>""')) then
     begin
      SetLastError(6, 'GFP_PROMO');
      SetControlText('GFP_PROMO','X') ;
      exit ;
     end else
     begin
      SetControlProperty('GFP_DEMARQUE','Enabled',False) ;
	    SetControlProperty('TGFP_DEMARQUE','Enabled',False) ;
	    SetControlProperty('BETABLIS','Enabled',False) ;
	    SetControlProperty('GFP_CASCADE','Enabled',False) ;
	    SetControlProperty('GFP_ARRONDI','Enabled',False) ;
	    SetControlProperty('GFP_DATEFIN','Enabled',False) ;
      SetField('GFP_ARRONDI','') ;
      SetField('GFP_DEMARQUE','') ;
      SetField('GFP_CASCADE','-') ;
      SetField('GFP_DATEFIN',iDate2099) ;
     end ;
   end else
   begin
   	SetControlProperty('GFP_DEMARQUE','Enabled',True) ;
	  SetControlProperty('TGFP_DEMARQUE','Enabled',True) ;
	  SetControlProperty('BETABLIS','Enabled',True) ;
	  SetControlProperty('GFP_CASCADE','Enabled',True) ;
	  SetControlProperty('GFP_ARRONDI','Enabled',True) ;
	  SetControlProperty('GFP_DATEFIN','Enabled',True) ;
   end ;
if F.FieldName='GFP_DATEDEBUT' then
   if GetField('GFP_DATEDEBUT') > GetField('GFP_DATEFIN') then SetField('GFP_DATEFIN',GetField('GFP_DATEDEBUT')) ;
if F.FieldName='GFP_DATEFIN' then
   if GetField('GFP_DATEDEBUT') > GetField('GFP_DATEFIN') then SetFocusControl('GFP_DATEFIN') ;
end ;


procedure TOM_TarifPer.OnNewRecord;
BEGIN
Inherited;
Setfield('GFP_DATEDEBUT',iDate1900) ;
Setfield('GFP_DATEFIN',iDate2099) ;
SetField('GFP_ETABLISSEMENT','...') ;
END;

procedure TOM_TarifPer.OnDeleteRecord ;
begin
Inherited;
if PeriodeUtiliseDansTarif then
begin SetLastError(5, 'GFP_CODEPERIODE'); exit ; end ;
SupprimePerEtab(GetField('GFP_CODEPERIODE')) ;
end;

function TOM_TarifPer.PeriodeUtiliseDansTarif : boolean ;
var  CodePeriode : string ;
Q: TQuery ;
begin
  Result:=False;
  CodePeriode:=GetField('GFP_CODEPERIODE');
  Q:=OpenSQL('Select GFM_TARFMODE from TARIFMODE Where GFM_PERTARIF="'+CodePeriode+'"',True);
  While not Q.EOF do
  begin
    if ExisteSQL('Select GF_Tarif from Tarif where GF_TARFMODE="'+Q.FindField('GFM_TARFMODE').AsString+'"') then
       begin
       Result:=True ;
       ferme(Q) ;
       Exit ;
       end
       else Q.next ;
  end ;
ferme(Q) ;
end ;

procedure TOM_TarifPer.SetLastError (Num : integer; ou : string );
begin
if ou<>'' then SetFocusControl(ou);
LastError:=Num;
LastErrorMsg:=TexteMessage[LastError];
end ;

Procedure TOM_TarifPer.ModifDateEtab ;
begin
AGLLanceFiche('MBO','PERETABLIS','','','CodePeriode='+GetField('GFP_CODEPERIODE')+';DateDebut='+DateToStr(GetField('GFP_DATEDEBUT'))	+';DateFin='+DateToStr(GetField('GFP_DATEFIN'))+';CodeArrondi='+GetField('GFP_ARRONDI')) ;
end ;

Procedure TOM_TarifPer.MiseAJourTarif ;
Var TarfMode,SQL,TypeRemise,SQLTarif,SQLEtab: String;
TOBTarif,TOBPerEtab,TOBEtab :TOB ;
QTarif,QPerEtab: TQuery ;
i: Integer ;
begin
TarfMode:=RechTarfMode(GetControlText('GFP_CODEPERIODE')) ;
if Tarfmode='' then exit ;
if GetField('GFP_Cascade')='X' then TypeRemise:='CAS' else TypeRemise:='MIE' ;
if (TarfMode<>'') and (ExisteSQL('Select GF_REMISE From tarif where GF_TARFMODE IN ('+TarfMode+')')) then
  if PGIAsk('Voulez-vous modifier les tarifs de cette période ?', Ecran.Caption) <> mrYes then exit 
// Si tarif par etablissement
  else begin
      TOBPerEtab:=TOB.Create('' ,nil,-1) ;
      SQLEtab:='Select GFP_ETABLISSEMENT from TARIFPER Where GFP_ETABLISSEMENT<>"..." And GFP_CODEPERIODE="'+GetField('GFP_CODEPERIODE')+'"' ;
      QPerEtab:=OpenSQL(SQLEtab,True) ;
      if Not QPerEtab.EOF then
      begin
      TOBPerEtab.LoadDetailDB('TARIFPER','','',QPerEtab,True,False) ;
      // TOBTarif pour traiter tous les cas
      TOBTarif:=TOB.Create('' ,nil,-1) ;
      SQLTarif:='Select GF_TARIF,GF_DEPOT,GF_DATEDEBUT,GF_DATEFIN,GF_DEMARQUE,GF_ARRONDI,GF_CASCADEREMISE from TARIF where GF_TARFMODE in ('+TarfMode+')' ;
      QTarif:=OpenSql(SQLTarif,True) ;
      if Not QTarif.EOF then TobTarif.LoadDetailDB('TARIF','','',QTarif,True,False) ;
      Ferme(QTarif) ;
      // Mise A jour des champs
      for i:=0 to TOBTarif.Detail.count-1 do
        begin
        TOBTarif.Detail[i].PutValue('GF_DEMARQUE',GetField('GFP_DEMARQUE'));
        TOBTarif.Detail[i].PutValue('GF_CASCADEREMISE',TypeRemise);
        TOBTarif.Detail[i].SetDateModif(NowH);
        TOBEtab:=TOBPerEtab.FindFirst(['GFP_ETABLISSEMENT'],[TOBTarif.Detail[i].GetValue('GF_DEPOT')],False) ;
        if TOBEtab=nil then
          begin
          TOBTarif.Detail[i].PutValue('GF_DATEDEBUT',GetField('GFP_DATEDEBUT'));
          TOBTarif.Detail[i].PutValue('GF_DATEFIN',GetField('GFP_DATEFIN'));
          TOBTarif.Detail[i].PutValue('GF_ARRONDI',GetField('GFP_ARRONDI'));
          end ;
        end ;
        TOBTarif.UpdateDB ;
      end else
      begin
      SQL:='UPDATE TARIF set GF_DEMARQUE="'+GetField('GFP_DEMARQUE')+'", GF_ARRONDI="'+GetField('GFP_ARRONDI')+'" , '
           +' GF_DATEDEBUT="'+UsDateTime(GetField('GFP_DATEDEBUT'))+'" , GF_DATEFIN="'+UsDateTime(GetField('GFP_DATEFIN'))+'" , '
           +' GF_CASCADEREMISE="'+TypeRemise+'", GF_DATEMODIF="'+UsTime(NowH)+'" where GF_TARFMODE IN ('+TarfMode+')' ;
      ExecuteSQL(SQL) ;
      end ;
      Ferme(QPerEtab) ;
    end ;
end ;

procedure SupprimePerEtab(CodePeriode: String) ;
Var SQL: String ;
begin
SQL :='DELETE FROM TARIFPER WHERE GFP_CODEPERIODE="'+CodePeriode+'"' ;
ExecuteSQL(SQL) ;
end ;


function RechTarfMode(periode:String ): String ;
var TobTarfMode : Tob ;
    Q : TQuery ;
    TarfMode : String ;
    i : Integer ;
begin
TobTarfMode:=TOB.Create('LesTarfMode',nil,-1) ;
Q:=OpenSQL('Select GFM_TARFMODE from TARIFMODE where GFM_PERTARIF="'+Periode+'"',True) ;
if Not Q.EOF then TobTarfMode.LoadDetailDB('_TarfMode','','',Q,False,False) ;
For i:=0 to TobTarfMode.Detail.Count-1 do
  begin
  if TarfMode='' then TarfMode:=TarfMode+'"'+IntToStr(TobTarfMode.Detail[i].GetValue('GFM_TARFMODE'))+'"'
     else TarfMode:=TarfMode+',"'+IntToStr(TobTarfMode.Detail[i].GetValue('GFM_TARFMODE'))+'"' ;
  end ;
ferme(Q) ;
Result:=TarfMode ;
TobTarfMode.Free;
end ;


// Modification des dates par établissement
procedure AGLModifDateEtab( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFFicheListe) then OM:=TFFicheListe(F).OM else exit;
  if (OM is TOM_TarifPer)
    then TOM_TarifPer(OM).ModifDateEtab
    else exit;
end;

Initialization
registerclasses([TOM_TarifPer]) ;
RegisterAglProc('ModifDateEtab',TRUE,0,AGLModifDateEtab);
end.
