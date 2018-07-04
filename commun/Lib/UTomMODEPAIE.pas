unit UTomMODEPAIE;

interface
uses  Classes,UTOM,SysUtils,Hent1,ParamSoc,HCtrls,LookUp,Forms,
      HMsgBox,
{$IFDEF EAGLCLIENT}
      UTob,eFichList,
{$ELSE}
      db,dbctrls,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Fe_Main,FichList,
{$ENDIF}
      Ent1 ;
Type
     TOM_MODEPAIE = Class (TOM)
       Private
       ControleChange : Boolean ;
       Function DansModRegl : Boolean ;
       Public
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnUpdateRecord ; override ;
       procedure OnLoadRecord ; override ;
       procedure OnDeleteRecord  ; override ;
       procedure OnArgument (Arguments : String ); override ;
       procedure OnClose  ; override ;

     END ;


     TOM_MODEPAIECOMPL = Class (TOM)
       procedure OnArgument (Arguments : String ); override ;
       procedure OnUpdateRecord  ; override ;
       procedure OnDeleteRecord  ; override ;
       procedure OnLoadRecord  ; override ;
       procedure OnChangeField (F : TField)  ; override ;
     Private
       ControleChange : Boolean ;
       Journal,compte : String ;
     END ;

const
	// libellés des messages
	TexteMessage: array[1..15] of string 	= (
          {1}        'Votre saisie est incorrecte : le montant maximum doit être positif.',
          {2}        'Vous devez renseigner un code d''acceptation.',
          {3}        'Votre mode de remplacement ne peut pas être identique au mode de paiement actif.' ,
          {4}        'Vous devez renseigner un compte général lettrable ou collectif.' ,
          {5}        'Vous ne pouvez pas supprimer ce mode de paiement car il est utilisé pour un règlement.',
          {6}        'Vous devez renseigner un mode de remplacement.',
          {7}        'Vous devez renseigner une catégorie de paiement ',
          {8}        'Vous devez renseigner une devise pour le Front Office',
          {9}        'Vous devez renseigner un type pour le Front Office',
          {10}       'Vous devez renseigner un établissement.',
          {11}       'Vous devez renseigner un libellé.',
          {12}       'Vous devez renseigner le journal. ',
          {13}       '',
          {14}       '',
          {15}       ''
                     );
implementation

procedure Tom_MODEPAIE.OnArgument(Arguments : String );
var x:integer ;
begin
Inherited;
x:=pos('ZOOM',Arguments);
SetControlChecked('IFZOOM',(x>0)) ;
{$IFDEF GESCOM}
SetControlVisible('PCOMPLEMENTFO',True) ;
SetControlProperty('PCOMPLEMENTFO','Caption','Ventes Comptoir');
SetControlProperty('MP_DEVISEFO','DataType','TTDEVISE');
{$ELSE}
SetControlVisible('PCOMPLEMENTFO',(ctxMode in V_PGI.PGIContexte)) ;
{$ENDIF}
SetActiveTabSheet('PGeneral');
SetControlVisible('BCOMPLEMENT',Not (ctxFo in V_PGI.PGIContexte)) ;
//uniquement en line
//SetControlVisible('MP_CONDITION',false) ;
//SetControlVisible('PCONDMT',false) ;
end;

procedure Tom_MODEPAIE.OnChangeField(F: TField);
Var StrCategorie,NatJournal,NatGeneral,StrType : String ;
    QQ, QQ1, QQ2 : TQuery ;
begin
Inherited;

if F.FieldName='MP_CATEGORIE' then
  begin
  StrCategorie:=GetControlText('MP_CATEGORIE');
  SetControlEnabled('TMP_CODEACCEPT',StrCategorie='LCR') ;
  SetControlEnabled('MP_CODEACCEPT',StrCategorie='LCR') ;
(*
  if ((StrCategorie = 'CB') or (StrCategorie = 'CHQ') or (StrCategorie = 'ESP')) then
     SetControlVisible('PComplement',(ctxGescom in VH_GC.GCContexte)) Else SetControlVisible('PComplement',FALSE);
*)
//  SetControlVisible('PGeneral',False) ;
//  SetControlVisible('PComplement',False) ;
  if Assigned(Ecran) and (Ecran is TFFicheListe) and (TFFicheListe(Ecran).TypeAction <> taConsult) then
    begin
    SetControlEnabled('MP_COPIECBDANSCTRL', (StrCategorie='CB'));
    SetControlEnabled('MP_AFFICHNUMCBUS', (StrCategorie='CB'));
    end;
  SetControlVisible('GBCB', ((StrCategorie='CHQ') or (StrCategorie='CB')));
  end;

if F.FieldName='MP_JALREGLE' then
  begin
  if not ControleChange then
    begin
    QQ:=OpenSQL('SELECT J_NATUREJAL, J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="'+ GetControlText('MP_JALREGLE') +'"',TRUE) ;
    if Not QQ.EOF then
       begin
       NatJournal:=QQ.FindField('J_NATUREJAL').AsString ;
       if ((NatJournal ='BQE') or (NatJournal = 'CAI')) then
          begin
          ControleChange :=true;
          SetControlEnabled('MP_CPTEREGLE',FALSE);
          SetField('MP_CPTEREGLE',(QQ.Findfield('J_CONTREPARTIE').AsString));
          end
       else
          begin
          if GetControl('MP_CPTEREGLE').Enabled = False then SetField('MP_CPTEREGLE','');
          SetControlEnabled('MP_CPTEREGLE',TRUE);
          end;
       end
    else SetControlEnabled('MP_CPTEREGLE',TRUE);
    ControleChange :=false;
    Ferme(QQ);
    end;
  end;

if F.FieldName='MP_CPTEREGLE' then
  begin
  if not ControleChange then
    begin
    QQ1:=OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+GetControlText('MP_CPTEREGLE')+'"',TRUE) ;
    if Not QQ1.EOF then
       begin
       NatGeneral:=QQ1.FindField('G_NATUREGENE').AsString ;
       if ((NatGeneral ='BQE') or (NatGeneral = 'CAI')) then
          begin
          QQ2:=OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE="'+GetControlText('MP_CPTEREGLE')+'"',TRUE) ;
          if Not QQ2.EOF then
             begin
             ControleChange :=true;
             SetField('MP_JALREGLE',(QQ2.Findfield('J_JOURNAL').AsString));
             end;
          Ferme(QQ2);
          end;
       end;
    ControleChange :=false;
    Ferme(QQ1);
    end;
  end;

if F.FieldName='MP_UTILFO' then
  begin
  if ctxMode in V_PGI.PGIContexte then
     begin
     if (GetField('MP_UTILFO')='X') and (GetField('MP_DEVISEFO')='') then SetField('MP_DEVISEFO',V_PGI.DevisePivot) ;
     end;
  end;
{$IFDEF GESCOM}
if F.FieldName='MP_UTILFO' then
  begin
  if (GetField('MP_UTILFO')='X') and (GetField('MP_DEVISEFO')='') then SetField('MP_DEVISEFO',V_PGI.DevisePivot) ;
  end;
{$ENDIF}

if F.FieldName='MP_TYPEMODEPAIE' then
  begin
  StrType:=GetField('MP_TYPEMODEPAIE');
  if Assigned(Ecran) and (Ecran is TFFicheListe) and (TFFicheListe(Ecran).TypeAction <> taConsult) then
    SetControlEnabled('MP_EDITCHEQUEFO', (StrType<>'005'));
  SetControlVisible('GBPERIPH', ((StrType='003') or (StrType='004') or (StrType='005')));
  end;
end;

procedure Tom_MODEPAIE.OnNewRecord;
begin
Inherited;
ControleChange:=false;
SetControlEnabled('TMP_CODEACCEPT',False) ;
SetControlEnabled('MP_CODEACCEPT',False) ;
end;

procedure Tom_MODEPAIE.OnUpdateRecord;
begin
Inherited;
if ((GetControlText('MP_GENERAL')<>'') and (not LookupValueExist(GetControl('MP_GENERAL'))))  then
  begin
  SetFocusControl('MP_GENERAL'); LastError:=4; LastErrorMsg:=TexteMessage[LastError]; exit;
  end;
if ((GetControlText('MP_CATEGORIE')=''))  then
  begin
  SetFocusControl('MP_GENERAL'); LastError:=7; LastErrorMsg:=TexteMessage[LastError]; exit;
  end;
if GetControlText('MP_CATEGORIE')='LCR' then
  begin
  if GetField('MP_CODEACCEPT')='' then
     begin
     SetFocusControl('MP_CODEACCEPT'); LastError:=2; LastErrorMsg:=TexteMessage[LastError]; exit;
     end;
  end ;
if GetField('MP_CONDITION')='X' then
  begin
  if ((Trim(GetControlText('MP_MONTANTMAX'))<>'') and (StrToFloat(GetControlText('MP_MONTANTMAX'))<0)) then
    begin
    SetFocusControl('MP_MONTANTMAX'); LastError:=1; LastErrorMsg:=TexteMessage[LastError]; exit;
    end;
  if GetField('MP_REMPLACEMAX')=GetField('MP_MODEPAIE') then
    begin
    SetFocusControl('MP_REMPLACEMAX'); LastError:=3; LastErrorMsg:=TexteMessage[LastError]; exit;
    end;
  if ((GetField('MP_REMPLACEMAX')='') and (GetField('MP_MONTANTMAX')<>0)) then
    begin
    SetFocusControl('MP_REMPLACEMAX'); LastError:=6; LastErrorMsg:=TexteMessage[LastError]; exit;
    end;
  end;

if not LookupValueExist(GetControl('MP_GENERAL')) then
   begin SetFocusControl('MP_GENERAL'); LastError:=4; LastErrorMsg:=TexteMessage[LastError]; exit; end;

if (ctxMode in V_PGI.PGIContexte) and (GetField('MP_UTILFO')='X') then
  begin
  if not LookupValueExist(GetControl('MP_DEVISEFO')) then
    begin
    SetFocusControl('MP_DEVISEFO'); LastError:=8; LastErrorMsg:=TexteMessage[LastError]; exit;
    end;
  if not LookupValueExist(GetControl('MP_TYPEMODEPAIE')) then
    begin
    SetFocusControl('MP_TYPEMODEPAIE'); LastError:=9; LastErrorMsg:=TexteMessage[LastError]; exit;
    end;
  end;

if GetField('MP_CATEGORIE')<>'LCR' then SetField('MP_CODEACCEPT','') ;
end;

Function Tom_MODEPAIE.DansModRegl : Boolean ;
Var SQl, ModePaiement : String ;
begin
ModePaiement:=GetControlText('MP_MODEPAIE');
SQL:='SELECT MR_MP1, MR_MP2, MR_MP3, MR_MP4, MR_MP5, MR_MP6, '
                   + 'MR_MP7, MR_MP8, MR_MP9, MR_MP10, MR_MP11, MR_MP12 '
                   + 'FROM MODEREGL ' ;
SQL:=SQL+' WHERE MR_MP1="'+ModePaiement+'" or MR_MP2="'+ModePaiement+'" ';
SQL:=SQL+' or MR_MP3="'+ModePaiement+'" or MR_MP4="'+ModePaiement+'" ';
SQL:=SQL+' or MR_MP5="'+ModePaiement+'" or MR_MP6="'+ModePaiement+'" ';
SQL:=SQL+' or MR_MP7="'+ModePaiement+'" or MR_MP8="'+ModePaiement+'" ';
SQL:=SQL+' or MR_MP9="'+ModePaiement+'" or MR_MP10="'+ModePaiement+'" ';
SQL:=SQL+' or MR_MP11="'+ModePaiement+'" or MR_MP12="'+ModePaiement+'" ';
Result:=ExisteSQL(SQL) ; Screen.Cursor:=SyncrDefault ;
end ;


procedure Tom_MODEPAIE.OnLoadRecord;
begin
Inherited;
ControleChange:=false;
end;

procedure Tom_MODEPAIE.OnDeleteRecord;
begin
if DansModRegl then BEGIN LastError:=5; LastErrorMsg:=TexteMessage[LastError]; exit ; END ;
inherited;
end;

procedure Tom_MODEPAIE.OnClose;
BEGIN
ChargeMPACC ;
inherited;
END ;

{=========================================================================================}
{============================= MODE PAIE COMPL ===========================================}
{=========================================================================================}
procedure Tom_MODEPAIECOMPL.OnArgument(Arguments : String );
var x:integer ;
critere,ChampMul,ValMul: String ;
begin
  Inherited;
  Repeat
    Critere:=Trim(ReadTokenSt(Arguments)) ;
    if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1) ;
         ValMul:=copy(Critere,x+1,length(Critere)) ;
         if ChampMul='JOURNAL' then Journal:=ValMul ;
         if ChampMul='COMPTE' then Compte:=ValMul ;
         end ;
      end ;
  until Critere='' ;
  if Assigned(GetControl('GREMBQE')) then
  begin
{$IFDEF GESCOM}
    SetControlProperty('GREMBQE', 'Visible', true);
{$ELSE}
    SetControlProperty('GREMBQE', 'Visible', false);
{$ENDIF GESCOM}
  end;
end;

procedure TOM_MODEPAIECOMPL.OnLoadRecord;
BEGIN
Inherited;
SetFocusControl('MPC_ETABLISSEMENT');
END ;

procedure TOM_MODEPAIECOMPL.OnUpdateRecord;
BEGIN
Inherited;
if GetControlText('MPC_ETABLISSEMENT')='' then
  BEGIN
  SetFocusControl('MPC_ETABLISSEMENT'); LastError:=10; LastErrorMsg:=TexteMessage[LastError]; exit;
  END;
if GetControlText('MPC_LIBELLE')='' then
  BEGIN
  SetFocusControl('MPC_LIBELLE'); LastError:=11; LastErrorMsg:=TexteMessage[LastError]; exit;
  END;
if GetControlText('MPC_JALREGLE')='' then //THValComboBox(GetControl('MPC_JALREGLE')).Value='' then
  BEGIN
  SetFocusControl('MPC_JALREGLE'); LastError:=12; LastErrorMsg:=TexteMessage[LastError]; exit;
  END;
END;

procedure TOM_MODEPAIECOMPL.OnDeleteRecord;
BEGIN
Inherited;
SetControlText('MPC_JALREGLE','') ;
END ;

procedure TOM_MODEPAIECOMPL.OnChangeField;
Var NatJournal,NatGeneral : String ;
    QQ, QQ1, QQ2 : TQuery ;
BEGIN
Inherited;
if (F.FieldName='MPC_ETABLISSEMENT') and (GetControlText('MPC_ETABLISSEMENT')<>'') and (GetControlText('MPC_JALREGLE')='')  then
  begin
  SetControlText('MPC_JALREGLE',Journal) ;
  SetControlText('MPC_CPTEREGLE',Compte) ;
  end ;
if F.FieldName='MPC_JALREGLE' then
  begin
  if not ControleChange then
    begin
    QQ:=OpenSQL('SELECT J_NATUREJAL, J_CONTREPARTIE FROM JOURNAL WHERE J_JOURNAL="'+ GetControlText('MPC_JALREGLE') +'"',TRUE) ;
    if Not QQ.EOF then
       begin
       NatJournal:=QQ.FindField('J_NATUREJAL').AsString ;
       if ((NatJournal ='BQE') or (NatJournal = 'CAI')) then
          begin
          ControleChange :=true;
          SetControlEnabled('MPC_CPTEREGLE',FALSE);
          SetField('MPC_CPTEREGLE',(QQ.Findfield('J_CONTREPARTIE').AsInteger));
          end
       else
          begin
          if GetControl('MPC_CPTEREGLE').Enabled = False then SetField('MPC_CPTEREGLE','');
          SetControlEnabled('MPC_CPTEREGLE',TRUE);
          end;
       end
    else SetControlEnabled('MPC_CPTEREGLE',TRUE);
    ControleChange :=false;
    Ferme(QQ);
    end;
  end;
if F.FieldName='MPC_CPTEREGLE' then
  begin
  if not ControleChange then
    begin
    QQ1:=OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="'+GetControlText('MPC_CPTEREGLE')+'"',TRUE) ;
    if Not QQ1.EOF then
       begin
       NatGeneral:=QQ1.FindField('G_NATUREGENE').AsString ;
       if ((NatGeneral ='BQE') or (NatGeneral = 'CAI')) then
          begin
          QQ2:=OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE="'+GetControlText('MPC_CPTEREGLE')+'"',TRUE) ;
          if Not QQ2.EOF then
             begin
             ControleChange :=true;
             SetField('MPC_JALREGLE',(QQ2.Findfield('J_JOURNAL').AsString));
             end;
          Ferme(QQ2);
          end;
       end;
    ControleChange :=false;
    Ferme(QQ1);
    end;
  end ;
END ;

Initialization
registerclasses([Tom_MODEPAIE]);
registerclasses([Tom_MODEPAIECOMPL]);
end.
