{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFMBPMUL ()
Mots clefs ... : TOF;QUFMBPMUL
*****************************************************************}
Unit QUFMBPMUL_TOF ;

Interface

Uses Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     mul,
     DBGrids,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$else}
     eMul,
     uTob,
     Grids,
{$ENDIF}
     HCtrls,
     HQry,
     UTOF,
     HTB97,
     HEnt1;


Type
  TOF_QUFMBPMUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure Traitement (Sender : TObject);
    private
     multi,ChpCode,ChpLib,ChpWhere:hString;
     function Fctretour:hString;
     procedure DblClick(Sender: TObject);
     procedure BtnFerme(Sender: TObject);   
     procedure BtnValider(Sender: TObject);

  end ;

Implementation

uses UUtil,
     BPBasic,
     Sysutils;

procedure TOF_QUFMBPMUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFMBPMUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFMBPMUL.OnUpdate ;
begin
  Inherited ;
end ;


procedure TOF_QUFMBPMUL.DblClick(Sender: TObject);
begin
  { EVI / Flipselection sur ligne courante si non sélectionné + close }
  {$IFDEF EAGLCLIENT} if not TFMul(Ecran).FListe.IsSelected(TFMul(Ecran).FListe.Row) then TFMul(Ecran).FListe.FlipSelection(TFMul(Ecran).FListe.Row);
  {$ELSE}if not TFMul(Ecran).FListe.IsCurrentSelected then TFMul(Ecran).FListe.FlipSelection;{$ENDIF}
  TToolBarButton97(getcontrol('BtnValider')).Click;
end;

procedure TOF_QUFMBPMUL.OnLoad ;
var codeCode,codeLib:hString;
begin
 Inherited ;
 codeCode:='';
 codeLib:='';
 if GetControlText( 'EDTCODE' )<>''
  then codeCode:=' AND '+ChpCode+' LIKE "'+GetControlText('EDTCODE')+'%" ';
 if GetControlText( 'EDTLIBELLE' )<>''
  then codeLib:=' AND '+ChpLib+' LIKE "'+GetControlText( 'EDTLIBELLE' )+'%" ';
 SetControlText('XX_WHERE',ChpWhere+CodeCode+CodeLib);
end ;

procedure TOF_QUFMBPMUL.OnArgument (S : String ) ;
var codeAxe,NomTable,Chp2,Val2,Liste:hString;
  Q: TQuery;
  sQl,Session,ValRestrict : hstring;
  i:integer;
begin
  Inherited ;
  TFMul(getcontrol('FLISTE')).OnDblClick := DblClick;
  TToolBarButton97(getcontrol('BTNFERME')).OnClick :=BtnFerme;
  TToolBarButton97(getcontrol('BTNVALIDER')).OnClick :=BtnValider;
  codeAxe:=TrouveArgument(S,'CODEAXE','');
  multi:=TrouveArgument(S,'MULTI','');
  Session:=TrouveArgument(S,'SESSION','');
  DonneNomTableChampsPrValeurAxe(codeaxe,NomTable,ChpCode,ChpLib,Chp2,Val2);
  Liste:='';
  SetControlText( 'XX_WHERE','');

  { EVI / Récupération des restrictions dans la définition de la session }
  Q := OpenSQL('SELECT QBS_CODEAXE1,QBS_CODEAXE2,QBS_CODEAXE3,QBS_CODEAXE4,'+
               'QBS_VALEURAXE1,QBS_VALEURAXE2,QBS_VALEURAXE3,QBS_VALEURAXE4 '+
               'FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+Session+'"', True, 1);
  try
    if not Q.Eof then
    begin
      for i:=0 to 3 do
      begin
        if ValRestrict = '' then
        begin
          if Q.Fields[i].AsString = codeAxe then ValRestrict := Q.Fields[i+4].AsString
        end;
      end;
    end;
  finally
    Ferme(Q);
  end;

  ValRestrict := StringReplace(ValRestrict,';','","',[rfReplaceAll, rfIgnoreCase]);


  if (Chp2 <> '') and  (Val2 <> '') then
  begin
    if ValRestrict <> '' then SetControlText( 'XX_WHERE',Chp2+'="'+Val2+'" AND ' + ChpCode+' IN ("'+ValRestrict+'")' )
    else SetControlText( 'XX_WHERE',Chp2+'="'+Val2+'"')
  end
  else
  begin
    if ValRestrict <> '' then SetControlText( 'XX_WHERE',ChpCode+' IN ("'+ValRestrict+'")' )
    else SetControlText( 'XX_WHERE','');
  end;

  ChpWhere:=GetControlText('XX_WHERE');

  sQl := 'SELECT QBX_LISTE FROM QBPBIBLIOAXE WHERE QBX_CODEAXE="' + CodeAxe +
         '" AND QBX_CODEMARCHE="'+CodeMarcheBiblioAxe+'"';
  Q := OpenSQL(Sql, True, 1);
  try
   if not Q.Eof then
    Liste := Q.FindField('QBX_LISTE').AsString;
  finally
  Ferme(Q);
 end;

 //TFMul(ecran).DBliste:=Liste;
 TFMul(ecran).Q.liste:=Liste;

 TFMul(ecran).Caption:=DonneLibelleCAxe('', '', CodeAxe);

 if multi='multi'
  then
   begin
    {$IFDEF EAGLCLIENT}
     { EVI / Suppression du goEditing --> multiselection impossible }
     TfMul(Ecran).FListe.Options:=[goColSizing,goRowSelect,goVertLine,goHorzLine,goRowSelect];
     TfMul(Ecran).FListe.MultiSelect := true;
    {$ELSE  EAGLCLIENT}
      { EVI / Suppression du dgEditing --> inutile }
      TfMul(Ecran).FListe.Options:=[dgtitles,dgColumnResize,dgEditing, dgRowSelect, dgColLines, dgRowLines, dgMultiSelect];
      TfMul(Ecran).FListe.MultiSelection := true;
    {$ENDIF EAGLCLIENT}
    TToolBarButton97(GetControl('BSelectAll')).visible:=true;
   end
  else
   begin
    {$IFDEF EAGLCLIENT}
      TfMul(Ecran).FListe.Options:=[goColSizing,goEditing,goRowSelect,goVertLine,goHorzLine];
      TfMul(Ecran).FListe.MultiSelect:=false;
    {$ELSE  EAGLCLIENT}
      TfMul(Ecran).FListe.Options:=[dgColumnResize,dgtitles,dgEditing,dgRowSelect,dgColLines,dgRowLines];
      TfMul(Ecran).FListe.MultiSelection:=false;
    {$ENDIF EAGLCLIENT}
    TToolBarButton97(GetControl('BSelectAll')).visible:=false;
   end;

 if TfMul(Ecran).Q <> NIL then TfMul(Ecran).DbListe  := Liste;
 UpdateCaption(TFMul(Ecran)) ;

// TToolBarButton97(GetControl('BCherche')).OnClick := Traitement;

end ;

procedure TOF_QUFMBPMUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFMBPMUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFMBPMUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFMBPMUL.Traitement(Sender: TObject);
begin
 if GetControlText( 'EDTCODE' )<>''
  then SetControlText( 'XX_WHERE',GetControlText('XX_WHERE')+
                           ' AND '+ChpCode+'="'+GetControlText( 'EDTCODE' )+'" ');
 if GetControlText( 'EDTLIBELLE' )<>''
  then SetControlText('XX_WHERE',GetControlText( 'XX_WHERE' )+
                           ' AND '+ChpLib+'="'+GetControlText( 'EDTLIBELLE' )+'" ');

end;

{ EVI / Gestion du retour du multicritère }
function TOF_QUFMBPMUL.Fctretour:hString;
var nbLigneSelect,i,iInd:integer;
    ret:hString;
begin
  ret:='';
  //-----------------> MULTISELECTION
  if multi='multi'
  then
  begin
    if BPOkOrli then
    //-----------------> ORLI
    begin
      if TFMul(ecran).FListe.allSelected
      then ret:=TraduireMemoire('<<Tous>>')
      else
      begin
        nbLigneSelect:=TFMul(ecran).FListe.nbSelected;
        if nbLigneSelect<1
        then ret:=TFMul(ecran).Q.FindField(ChpCode).AsString;

        for i:=0 to nbLigneSelect-1 do
        begin
          TFMul(ecran).FListe.GotoLeBOOKMARK(i);
          if ret<>''
          then ret:=ret+';'+TFMul(ecran).Q.FindField(ChpCode).AsString
          else ret:=ret+TFMul(ecran).Q.FindField(ChpCode).AsString;
        end;
      end;
    end
   //ORLI <-----------------
    else
    begin
      //-----------------> SELECTION TOTALE
      if TFMul(ecran).FListe.AllSelected then
      begin
          {$IFDEF EAGLCLIENT}TFMul(Ecran).Fetchlestous;{$ENDIF}
          TFMul(Ecran).Q.First;
          while not TFMul(ecran).Q.Eof do
          begin
            if ret<>'' then
            {$IFDEF EAGLCLIENT}
            ret := ret + ';' + TFMul(ecran).Q.TQ.FindField(ChpCode).AsString
            else
            ret := ret + TFMul(ecran).Q.TQ.FindField(ChpCode).AsString;
            {$ELSE}
            ret := ret + ';' + TFMul(ecran).Q.FindField(ChpCode).AsString
            else
            ret := ret + TFMul(ecran).Q.FindField(ChpCode).AsString;
            {$ENDIF}
            TFMul(Ecran).Q.next;
          end;
        //ret := TraduireMemoire('<<Tous>>');
      end
      //<----------------- SELECTION TOTALE
      else
      begin
        //-----------------> SELECTION PARTIELLE
        for iInd := 0 to TFMul(ecran).FListe.NbSelected -1 do
        begin
          TFMul(ecran).FListe.GotoLeBookMark(iInd);
          {$IFDEF EAGLCLIENT}
          TFMul(ecran).Q.TQ.Seek (TFMul(ecran).FListe.Row-1) ;
          if ret<>'' then
          ret := ret + ';' + TFMul(ecran).Q.TQ.FindField(ChpCode).AsString
          else
          ret := ret + TFMul(ecran).Q.TQ.FindField(ChpCode).AsString;
          {$ELSE}
          if ret<>'' then
          ret := ret + ';' + TFMul(ecran).Q.FindField(ChpCode).AsString
          else
          ret := ret + TFMul(ecran).Q.FindField(ChpCode).AsString;
          {$ENDIF}
        end;
      end;
    end;
  //<----------------- MULTISELECTION
  end
  else
  //<----------------- UNISELECTION
  begin
    {$IFDEF EAGLCLIENT}
    TFMul(ecran).FListe.GotoRow(TFMul(ecran).FListe.Row);
    TFMul(ecran).Q.TQ.Seek (TFMul(ecran).FListe.Row-1) ;
    ret := TFMul(ecran).Q.TQ.FindField(ChpCode).AsString
    {$ELSE}
    ret:=TFMul(ecran).Q.FindField(ChpCode).AsString;
    {$ENDIF}
  end;
  //<----------------- UNISELECTION

 result:=ret;
end;

procedure TOF_QUFMBPMUL.BtnFerme(Sender: TObject);
begin
 TfMul(Ecran).retour:='';
 TToolBarButton97(getcontrol('Bannuler')).Click;
end;

procedure TOF_QUFMBPMUL.BtnValider(Sender: TObject);
begin
 TfMul(Ecran).retour:=Fctretour();
 TToolBarButton97(getcontrol('Bannuler')).Click;
end;

Initialization
  registerclasses ( [ TOF_QUFMBPMUL ] ) ; 
end.
