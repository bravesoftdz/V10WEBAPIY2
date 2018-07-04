{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/04/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AFPARAMFORMULE (AFPARAMFORMULE)
Mots clefs ... : TOM;AFPARAMFORMULE
*****************************************************************}
Unit uTomAfRevParamFormule;

Interface
      
Uses HEnt1,
     StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MainEagl,eFiche,eFichList,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,Fiche,FichList,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HMsgBox,
     UTOM,
     UTob,HTB97,uTofAfRevFormuleEdit,dialogs,M3FP,dbctrls, dicobtp;


Type
  TOM_AFPARAMFORMULE = Class (TOM)
    private     NbIndice :integer ;
    Anomalie : boolean ;
    Qindice : Tquery ;
    TobIndice : TOB  ;
    LaGrille : THGrid ;
    BTNRECOPIE : TToolbarButton97 ;
    BTNFormule: TToolbarButton97 ;
    
    //BTNFormuleEdition: TToolbarButton97 ;
   // BTNValider :  TToolbarButton97 ;

    public
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    Procedure RempliColonneFixe ;
    Procedure RempliColonneIndice ;
    Procedure RempliColonneValeurIndice ;
    Procedure RempliColonnePublication ;
    Procedure RempliColonneCodeIndiceAffaire ;
    Procedure SauveColonneValeurIndice ;
    Procedure SauveColonnePublication ;
    Procedure SauveColonneCodeIndiceAffaire ;
    function FabriqueTobIndices : integer ;
    procedure BtnRecopieClick(sender : Tobject) ;
//    procedure BTNFormuleEditionClick(sender : Tobject) ;
    procedure BTNFormuleClick(sender : Tobject) ;
//    procedure BTNValiderClick(sender : Tobject) ;
    procedure LaGrilleClick(sender : Tobject) ;
    Function  OkPourValidation : Boolean ;
    procedure Recopie ;
    end ;

  procedure AFLanceFiche_ParamFormule(range,cle,Action : string ) ;
                        
  Const
  ColTitre  = 0;
  ColIndice = 1 ;
  ColIndAff  =2 ;
  ColValInit =3 ;
  ColPub     =4 ;

                           
  TexteMessage: array[1..7] of string 	= (
  {1}  'Les codes des indices dans l''affaire doivent être remplis.',
  {2}  'Veuillez saisir une date initiale d''indices ou bien pour chaque indice une valeur initiale ou une publication.',
  {3}  'Le mois de lecture ne peut pas etre zéro ! ',
  {4}  'Jour de lecture incorrect ! ',
  {5}  'La premiere date de révision doit être renseignée! ',
  {6}  'La prochaine date appliquée doit être supérieure à la dernière date appliquée! ',
  {7}  'La prochaine date appliquée doit être supérieure à la dernière date appliquée.'
  );


Implementation
uses uTofAfRevFormule ;

procedure TOM_AFPARAMFORMULE.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_AFPARAMFORMULE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_AFPARAMFORMULE.OnUpdateRecord ;
begin
  if OkPourValidation then
    begin
    Inherited ;
    Setfield('AFC_PARAMFORMULEOK','X');
    SauveColonneValeurIndice ;
    SauveColonnePublication ;
    SauveColonneCodeIndiceAffaire ;
    Anomalie:=false ;
    end else
    begin
    Setfield('AFC_PARAMFORMULEOK','-');
    Anomalie:=True ;
    end ;
end ;

procedure TOM_AFPARAMFORMULE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

Procedure TOM_AFPARAMFORMULE.RempliColonneIndice ;
var j : integer ;
begin
 for j:=1 to 10 do
   begin
   LaGrille.Cells[ColIndice,j]:= TobIndice.Detail[0].GetValue('AFE_INDCODE'+inttostr(j));
   if (LaGrille.Cells[ColIndice,j]<>'') then NbIndice:=j ;
   end ;
end ;


Function TOM_AFPARAMFORMULE.OkPourValidation : Boolean ;
var i,MessageErreur,Mois,Jour,Code : integer ;
begin
  result:=True ;
  MessageErreur:=0 ;
  for i:=1 to NbIndice do
    begin
     if (LaGrille.Cells[ColIndAff,i]='') then
       begin
       result:=False ;
       MessageErreur:=1 ;
       end ;
    end ;
  if result then begin
    if (getcontroltext('AFC_DATEINITIALE')='') or (getcontroltext('AFC_DATEINITIALE')='31/12/2099') then
      begin
      for i:=1 to NbIndice do
        begin
         if (valeur(LaGrille.Cells[ColValInit,i])=0) then 
           begin
           result:=False ;
           MessageErreur:=2 ;
           end ;
        end ;
      end
      else
      begin
      result:=True ;
      end ;
    end ;
  if result then
    begin
    val(getcontroltext('AFC_NUMMOISREV'),Mois,code) ;
    if Mois=0 then
      begin
      result:=False ;
      MessageErreur:=3 ;
      end ;
    end ;
  if result then
    begin
    if not (TdbCheckBox(getControl('AFC_REVFINMOIS')).Checked) then
      begin
      val(getcontroltext('AFC_NUMJOURREV'),Jour,Code) ;
      if Jour=0 then
        begin
        result:=False ;
        MessageErreur:=4 ;
        end ;
      end ;
    end ;

  if result then
    if (getcontroltext('AFC_PREMIEREDATE')='') or (getcontroltext('AFC_PREMIEREDATE')='31/12/2099') then
    begin
      result:=False ;
      MessageErreur:=5 ;
    end ;

  if result then
    if (StrToDate(getcontroltext('AFC_LASTDATEAPP')) <> iDate2099) and
       (StrTodate(getcontroltext('AFC_NEXTDATEAPP'))<strtodate(getcontroltext('AFC_LASTDATEAPP'))) then
    begin
      result := False;
      MessageErreur :=6;
    end ;

  if result then
    if (StrTodate(getcontroltext('AFC_LASTDATECALC')) = strtodate(getcontroltext('AFC_LASTDATEAPP'))) and
       (StrTodate(getcontroltext('AFC_NEXTDATEAPP')) <= strtodate(getcontroltext('AFC_LASTDATECALC'))) and
       (StrTodate(getcontroltext('AFC_NEXTDATEAPP')) <> iDate2099) then
    begin                                       
      result := False;
      MessageErreur := 7;
    end;

  if result = False then
    PGIBoxAF(TexteMessage[MessageErreur], Ecran.Caption);
end ;

Procedure TOM_AFPARAMFORMULE.SauveColonneCodeIndiceAffaire ;
var
  j   : integer;
  st  : string;

begin
  for j:=1 to 10 do
  begin
    if j>NbIndice then
      st := ''
    else
      st := LaGrille.Cells[ColIndAff,j];
    Setfield('AFC_INDAFF'+inttostr(j), st);
  end;
end;

Procedure TOM_AFPARAMFORMULE.SauveColonneValeurIndice ;
var j : integer;
begin
  for j:=1 to NbIndice do
    Setfield('AFC_VALINITIND'+inttostr(j),valeur(LaGrille.Cells[ColValInit,j]));
end;

Procedure TOM_AFPARAMFORMULE.SauveColonnePublication ;
var j : integer;
begin
  for j:=1 to NbIndice do
    if LaGrille.Cells[ColPub,j] <> '' then
      Setfield('AFC_PUBCODE'+inttostr(j), LaGrille.CellValues[ColPub,j]);
End;

Procedure TOM_AFPARAMFORMULE.RempliColonneFixe ;
var j : integer ;
begin
  for j:=1 to 10 do
   LaGrille.Cells[ColTitre,j]:= 'Indice '+inttostr(j);
end ;



Procedure TOM_AFPARAMFORMULE.RempliColonneCodeIndiceAffaire ;
var j : integer ;
begin
  for j:=1 to 10 do
    begin
    LaGrille.Cells[ColIndAff,j]:= Getfield('AFC_INDAFF'+inttostr(j));
    if LaGrille.Cells[ColIndAff,j]='' then
      LaGrille.Cells[ColIndAff,j]:=LaGrille.Cells[ColIndice,j] ;
    end ;
end ;


Procedure TOM_AFPARAMFORMULE.RempliColonneValeurIndice ;
var j : integer ;
begin
  for j:=1 to 10 do
    if j <=NbIndice then
      LaGrille.Cells[ColValInit,j]:= Getfield('AFC_VALINITIND'+inttostr(j))
    else
      LaGrille.Cells[ColValInit,j]:= '' ;
end ;

Procedure TOM_AFPARAMFORMULE.RempliColonnePublication ;
var j : integer ;
begin
  for j:=1 to 10 do
    if Getfield('AFC_PUBCODE'+inttostr(j)) <> '' then
     LaGrille.Cells[ColPub,j]:= RechDom('AFPUBLICATION_LIB', Getfield('AFC_PUBCODE'+inttostr(j)),False);
end;
    
function TOM_AFPARAMFORMULE.FabriqueTobIndices : integer ;
Var StQindice : String ;
begin
  StQindice:='SELECT AFE_FORCODE,AFE_INDCODE10, AFE_INDCODE9, AFE_INDCODE8,AFE_INDCODE7, AFE_INDCODE6,' ;
  StQindice:=StQindice+'  AFE_INDCODE5,AFE_INDCODE4,AFE_INDCODE3,AFE_INDCODE2,AFE_INDCODE1 ' ;
  StQindice:=StQindice+' FROM AFFORMULE WHERE AFE_FORCODE = "'+getfield('AFC_FORCODE')+'"';
  QIndice:=Nil ;
  try
    QIndice := OpenSQL(StQindice, TRUE);
    TobIndice.LoadDetailDB('','','',QIndice,false) ;
    result:=TobIndice.detail.count ;
  finally
    Ferme(QIndice);
  end
end;

procedure TOM_AFPARAMFORMULE.OnLoadRecord ;
begin

  Inherited ;
  TobIndice:=TOB.Create('Mes Indices',nil,-1);
  if FabriqueTobIndices>0 then Begin
    RempliColonneIndice ;
    TobIndice.free ;
    end ;
  RempliColonneFixe ;
  RempliColonneValeurIndice ;
  RempliColonnePublication ;
  RempliColonneCodeIndiceAffaire ;
  forceUpdate ;
end ;

procedure TOM_AFPARAMFORMULE.Recopie ;
var j : integer ;
begin
 for j:=1 to 10 do LaGrille.Cells[ColIndAff,j]:=LaGrille.Cells[ColIndice,j] ;
 forceUpdate ;
end;

procedure TOM_AFPARAMFORMULE.BtnRecopieClick(sender : Tobject) ;

begin
Recopie ;
end ;

procedure TOM_AFPARAMFORMULE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_AFPARAMFORMULE.LaGrilleClick(sender : Tobject) ;
begin
  forceUpdate ;
end ;

//procedure TOM_AFPARAMFORMULE.BTNValiderClick (sender : Tobject) ;
//begin

//end ;

{procedure TOM_AFPARAMFORMULE.BTNFormuleEditionClick(sender : Tobject) ;
begin
//BTNValider.Click ;
if Not anomalie then
  AglLanceFicheAFREVFORMULEEDIT('','FORMULE='+getfield('AFC_FORCODE')+';Affaire='+getfield('AFC_AFFAIRE')) ;
end ;
}
 
procedure TOM_AFPARAMFORMULE.BTNFormuleClick(sender : Tobject) ;
Var indice : String ;
begin
indice:=GetField('AFC_FORCODE') ;
AFLanceFiche_Formule('','AFE_FORCODE='+indice+';ACTION=CONSULTATION') ;
end ;

procedure TOM_AFPARAMFORMULE.OnArgument ( S: String ) ;
begin
  Inherited ;
  LaGrille:=THGrid(Getcontrol('LaGrille')) ;
  BtnRecopie:=TToolbarButton97(Getcontrol('BTNRECOPIE')) ;
  BtnRecopie.OnClick:=BtnRecopieClick ;

//  BTNFormuleEdition:=TToolbarButton97(Getcontrol('BTNFORMULEEDITION')) ;
//  BTNFormuleEdition.OnClick:=BTNFormuleEditionClick ;
//  BTNValider:=TToolbarButton97(Getcontrol('BVALIDER')) ;
//  BTNValider.OnClick:=BTNValiderClick ;

  BTNFormule:=TToolbarButton97(Getcontrol('BTNFORMULE')) ;
  BTNFormule.OnClick:=BTNFormuleClick ;
  LaGrille.OnClick:=LaGrilleClick ;

  LaGrille.ColEditables[1] := False;
  LaGrille.ColTypes[3]:='R';

  LaGrille.ColFormats[ColPub] := 'CB=AFPUBLICATION_LIB||<<Aucun>>';

  SetFocusControl('AFC_PREMIEREDATE');
end ;

procedure TOM_AFPARAMFORMULE.OnClose ;
begin
  Inherited ;
  if Anomalie then begin
    lastError:=-1 ;
    Anomalie:=False ;
  end
  Else
  begin
  lastError:=0 ;
  end ;
end ;

procedure TOM_AFPARAMFORMULE.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure AFLanceFiche_ParamFormule(range,cle,Action : string ) ;
begin
  AglLanceFiche ('AFF','AFREVPARAMFORMULE',range,cle,Action);
end ;

Initialization
  registerclasses ( [ TOM_AFPARAMFORMULE ] ) ;
end.
