{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 09/02/2004
Modifié le ... : 28/05/2007
Description .. : Passage en eAGL
Suite ........ : 
Suite ........ : JP 28/05/07 : FQ 17460 : Gestion des tic / tid
Mots clefs ... : 
*****************************************************************}
unit TofVerifRib;

interface

uses SysUtils, WinTypes, WinProcs, Classes, Graphics, Forms, Controls, Buttons,Dialogs,
     StdCtrls, ExtCtrls, Hctrls, Grids, Mask,
{$IFDEF EAGLCLIENT}
  MaineAGL,
  UtileAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DB, DBGrids, DBCtrls,
  HDB, PrintDBG, Fe_Main,
{$IFNDEF IMP}
     SaisBor,
{$ENDIF}
{$ENDIF}
     Ent1, Spin, HmsgBox, Hqry, HEnt1, ComCtrls,
     Messages, HSysMenu, HPanel, UiUtil,
     HTB97,Hcompte,HXLSPAS,
     Hstatus,UtilPgi,UTOF,UTOB,
     Vierge,
     Saisie,
{$IFNDEF CCS3}
     EcheMPA,
{$ENDIF}
     SaisUtil,SaisComm,UtilDiv;

procedure CPLanceFiche_VerifRib(psz : String);

type
  TOF_VerifRib = class(TOF)
    TheTob : TOB;
    G: THGrid;
    private
      TheWhere, gszRIB : string ;
      LibTiersOK,UpdateDBOk,CorAffOK, gbIBAN : Boolean ;
      Nb,NbCorrect : integer ;
      NBInco,NbCorr : THEdit ;
      procedure LanceVerifRib (Where : string) ;
      procedure AjoutRibErreur (Q : TQuery;MsgErr : string) ;
      function RecupLeRib(var Rib, Aux : string ) : Boolean; 
      function MsgErr(Err : integer) : string ;
      function MAJLigneEcriture(Rib : String;Pos : integer) : string ;
      procedure MAJGrid(Tous : Boolean);
      procedure OnGridDblClick(Sender : TObject);
      procedure OnImprimerClick(Sender : TObject);
      procedure OnZoomClick(Sender: TObject);
      procedure OnAffCorClick(Sender: TObject);
      procedure AfficheNombre ;
      procedure VerifLibelleTiers ;
    public
      procedure Onload ;  override ;
      procedure OnUpdate ;override ;
      procedure OnClose ; override ;
      procedure OnArgument (S : String ) ; override ;
    end;
const
    Msg : array[0..9] of string = (
    {00}       'RIB Correct'
    {01}       ,'RIB non renseigné'
    {02,03,04} ,'Clé erronée','',''
    {05}       ,'établissement'
    {06}       ,'guichet'
    {07}       ,'compte'
    {08}       ,'clé'
    {09}       ,'RIB inexistant' ) ;


    CL_JAl      = 0 ;
    CL_DATE     = Cl_JAL+1 ;
    Cl_NPIECE   = CL_DATE+1 ;
    CL_NLIGNE   = CL_NPIECE+1 ;
    CL_AUXILI   = Cl_NLIGNE+1 ;
    CL_DEBIT    = Cl_AUXILI+1 ;
    CL_CREDIT   = CL_DEBIT+1 ;
    CL_ERREUR   = CL_CREDIT+1 ;
    CL_RIB      = CL_ERREUR+1 ;
//    CL_RIB      = CL_ERREUR+1 ;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibEcriture ;

procedure CPLanceFiche_VerifRib(psz : String);
begin
  AGLLanceFiche('CP','VERIFRIB','','',psz);
end;

function IsIncorrect(Chaine : String) : boolean;
var i : integer ;
begin
Result:=False;
for i:=1 to Length(Chaine) do
  if (not (Chaine[i] in ['0'..'9'])) and (not (Chaine[i] in ['A'..'Z'])) then begin
     Result:=True;
     Exit;
  end;
end ;

function VerificationRib(Rib : string) : integer ;
var Etab,Guichet,Compte,Cle,Domiciliation,StTmp : string; i : integer ;
begin
StTmp:=Rib ;
if Rib='' then begin Result:=-1 ; Exit ; end ;
for i:=1 to length(Rib) do if rib[i]<>' ' then break ;
if i=length(Rib)+1 then begin Result:=-1 ; Exit ; end ;
if pos('/',StTmp)<>6  then begin Result:=-15; Exit; end ; System.Delete(StTmp,1,6) ;
if pos('/',StTmp)<>6  then begin Result:=-16; Exit; end ; System.Delete(StTmp,1,6) ;
if pos('/',StTmp)<>12 then begin Result:=-17; Exit; end ; System.Delete(StTmp,1,12) ;
if pos('/',StTmp)<>3  then begin Result:=-18; Exit; end ;
DecodeRib(Etab,Guichet,Compte,Cle,Domiciliation,Rib) ;
if IsIncorrect(Etab)    then begin Result:=-25 ; Exit ; end ;
if IsIncorrect(Guichet) then begin Result:=-26 ; Exit ; end ;
if IsIncorrect(Compte)  then begin Result:=-27 ; Exit ; end ;
if IsIncorrect(Cle)     then begin Result:=-28 ; Exit ; end ;
if VerifRib(Etab,Guichet,Compte)<>Cle then begin Result:=-2; Exit ; end;
Result:=0 ;
end ;

function VerificationIban(Iban : string) : integer ;
begin
if Iban='' then begin Result:=-1 ; Exit ; end ;
if (VH^.PaysLocalisation=CodeISOES) and (RIBestIBAN(iban)) then
   delete(Iban,1,1) ;
if ErreurDansIban(Iban) then begin Result:=-2; Exit ; end;
Result:=0 ;
end;

procedure TOF_VerifRib.Onload ;
var
  BImprim,
  BZoom,
  BAffCor : TButton ;
begin
If Ecran <> Nil Then {JP 19/08/04 : Il ne s'agit pas d'un QRS1}
  TFVierge(Ecran).HelpContext := 999999211;
UpdateDBOk:=True ; LibTiersOK:=False ;  CorAffOK:=True ;
G:=THGrid(GetControl('FLISTE')) ; if G=nil then Exit ;
BImprim:=TButton(GetControl('BIMPRIMER')); {JP 19/08/04 : Un seul bouton imprimer me semble suffisant}
if (BImprim<>nil) and (not assigned(BImprim.OnClick)) then BImprim.OnClick:=OnImprimerClick ;
BZoom:=TButton(GetControl('BZOOM')) ;
if (BZoom<>nil) and (not assigned(BZoom.OnClick)) then BZoom.OnClick:=OnZoomClick ;
BAffCor:=TButton(GetControl('BAFFCOR')) ;
if (BAffCor<>nil) and (not assigned(BAffCor.OnClick)) then BAffCor.OnClick:=OnAffCorClick ;
OnAffCorClick(BAffCor) ;
G.RowCount:=2 ; if not(assigned(G.OnDblClick)) then G.OnDblClick:=OnGridDblClick ;
NBInco:=THEdit(GetControl('LNBINCORRECT')) ;
if NbInco<>nil then NbInco.Text:= gszRIB+' incorrects : 0' ;
NbCorr:=THEdit(GetControl('LNBCORRECT')) ;
if NbCorr<>nil then NbCorr.Text:=gszRIB+' corrects :  0' ;
G.ColAligns[CL_NPIECE]:=taCenter ;
G.ColAligns[CL_NLIGNE]:=taCenter ;
G.ColAligns[CL_DEBIT]:=taRightJustify ;
G.ColAligns[CL_CREDIT]:=taRightJustify ;
G.ColWidths[CL_RIB]:=0 ;
TheTob:=Tob.Create('§ECRITURE',nil,-1) ;
TheWhere:=TrouveArgument(TFVierge(Ecran).FArgument,'WHERE=') ;
if Thewhere='' then TheWhere:=' WHERE E_AUXILIAIRE<>""' else  UpdateDBOk:=False ;
if Not UpdateDBOk then VerifLibelleTiers ;
LanceVerifRib(TheWhere) ;
if Nb=0 then begin
  BImprim.Enabled:=False ;
  BZoom.Enabled:=False ;
  BAffCor.Enabled:=False ;
  TButton(GetControl('BValider')).enabled:=False ;
  end ;
end ;

procedure TOF_VerifRib.OnUpdate ;
var
  i : Integer ;
begin
  {JP 19/08/04 : FQ 13694 : réactivation du bouton BVALIDER}
  for i := 1 to G.RowCount - 1 do
    if G.Cells[CL_ERREUR, i] <> Msg[0] then begin
      G.Row := i;
      MajGrid(True);
    end;
end;

procedure TOF_VerifRib.OnClose ;
begin
TheTob.Free ; TheTob:=nil ;
end ;

procedure TOF_VerifRib.LanceVerifRib(Where : string) ;
var
  Q : TQuery ;
  Err : Integer ;
  Cols, St : string ;
  lBoRibExist : Boolean ;
begin
TheTob.ClearDetail ;

 Cols := 'E_RIB, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, '
       + 'E_AUXILIAIRE, E_DEBIT, E_CREDIT, E_EXERCICE, E_NUMECHE, E_QUALIFPIECE, '
       + 'E_MODESAISIE, T_LIBELLE, E_GENERAL' ;

St:='SELECT '+Cols+' FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' ;
St:=St+' LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ' ;

Q:=OpenSql(St+Where,True) ;
Nb:=0 ;
NbCorrect:=0 ;

While not Q.eof do
  begin
  Nb:=Nb+1 ;
  if (gbIBAN)  or ((VH^.PaysLocalisation=CodeISOES) and (RIBestIBAN(Q.FindField('E_RIB').AsString))) then
    Err:=VerificationIban(Q.FindField('E_RIB').AsString)
  else
    Err:=VerificationRib(Q.FindField('E_RIB').AsString) ;

  // Ajout Test existance du RIB sur le compte FQ 17062 SBO 23/11/2005
  if (Err=0) and not VH^.CPIFDEFCEGID then  {JP 21/12/05 : Cegid ne veut pas en entendre parler du moins tel quel !!!}
    begin
    if Q.FindField('E_AUXILIAIRE').AsString <> ''
      then lBoRibExist := ExisteRibSurCpt( Q.FindField('E_AUXILIAIRE').AsString, Q.FindField('E_RIB').AsString )
      else lBoRibExist := ExisteRibSurCpt( Q.FindField('E_GENERAL').AsString,    Q.FindField('E_RIB').AsString ) ;
    if not lBoRibExist then Err := -9 ;
    end ;

  AjoutRibErreur(Q,MsgErr(Err));
  if Err=0 then NbCorrect:=NbCorrect+1 ;
  Q.Next ;
  AfficheNombre ;
  end ;
ferme(Q) ;
if G.RowCount-1>1 then G.RowCount:=G.RowCount-1 ;
end ;

procedure TOF_VerifRib.AfficheNombre ;
begin
if NbInco<>nil then NbInco.Text:=gszRIB+' incorrects : '+IntToStr(Nb-NbCorrect) ;
if NbCorr<>nil then NbCorr.Text:=gszRIB+' corrects : '+IntToStr(NbCorrect) ;
Application.ProcessMessages ;
end;

procedure TOF_VerifRib.VerifLibelleTiers ;
var Where : string ;
begin
Where:=UpperCase(TheWhere) ;
if pos('JOIN TIERS',Where)<>0 then LibTiersOK:=True ;
end;

procedure TOF_VerifRib.AjoutRibErreur (Q : TQuery;MsgErr : string) ;
var lg : integer ; T1 : Tob ;
begin
lg:=G.RowCount-1 ;
G.RowCount:=G.RowCount+1 ;
G.Cells[CL_JAL,lg]:=Q.FindField('E_JOURNAL').AsString ;
G.Cells[CL_DATE,lg]:=Q.FindField('E_DATECOMPTABLE').AsString ;
G.Cells[Cl_NPIECE,lg]:=Q.FindField('E_NUMEROPIECE').AsString ;
G.Cells[CL_NLIGNE,lg]:=Q.FindField('E_NUMLIGNE').AsString ;
if LibTiersOK then G.Cells[CL_AUXILI,lg]:=Q.FindField('T_LIBELLE').AsString
              else G.Cells[CL_AUXILI,lg]:=Q.FindField('E_AUXILIAIRE').AsString ;
G.Cells[CL_DEBIT,lg] :=format('%*.*f',[20,V_PGI.OkDecV,StrToFloat(Q.FindField('E_DEBIT').AsString) ]); ;
G.Cells[CL_CREDIT,lg]:=format('%*.*f',[20,V_PGI.OkDecV,StrToFloat(Q.FindField('E_CREDIT').AsString)]); ;
G.Cells[CL_ERREUR,lg]:=MsgErr ;
G.Cells[CL_RIB,lg]:=Q.FindField('E_RIB').AsString;
T1:=TOB.Create('ECRITURE',TheTob,-1) ;
T1.SelectDB('',Q)  ;
end ;

function TOF_VerifRib.MsgErr(Err : integer) : string ;
var ErrPositif : integer ;
begin
if Err>0 then begin Result:='Erreur '+gszRIB ; Exit ; end ;
ErrPositif:=Err*-1 ;
if ErrPositif<10 then Result:=Msg[ErrPositif]
else if ErrPositif<20 then  Result:='Longueur '+Msg[ErrPositif-10]+' incorrect'
     else if ErrPositif<30 then Result:='Caractère incorrect dans '+Msg[ErrPositif-20]
          else Result:='Erreur '+gszRIB ;
end ;

function TOF_VerifRib.MAJLigneEcriture(Rib : String;Pos : integer) : string ;
Var St : String ;
begin
if TheTob=nil then Exit ;
TheTob.Detail[G.Row-1].PutValue('E_RIB',Rib) ;
if UpdateDBOK Then TheTob.Detail[G.Row-1].UpdateDB(True) Else
  begin
  {JP 16/11/07 : FQ 21847 : Gestion de E_UTILISATEUR}
  St:='UPDATE ECRITURE SET E_RIB = "' + Rib + '", E_UTILISATEUR = "' + V_PGI.User +
      '", E_DATEMODIF = "' + UsTime(NowH) + '" WHERE ' ;
  St:=St+'E_JOURNAL="'+TheTob.Detail[G.Row-1].GetValue('E_JOURNAL')+'" ' ;
  St:=St+'AND E_EXERCICE="'+TheTob.Detail[G.Row-1].GetValue('E_EXERCICE')+'" '  ;
  St:=St+'AND E_DATECOMPTABLE="'+UsDateTime(TheTob.Detail[G.Row-1].GetValue('E_DATECOMPTABLE'))+'" ' ;
  St:=St+'AND E_NUMEROPIECE='+IntToStr(TheTob.Detail[G.Row-1].GetValue('E_NUMEROPIECE'));
  St:=St+' AND E_NUMLIGNE='+IntToStr(TheTob.Detail[G.Row-1].GetValue('E_NUMLIGNE'));
  St:=St+' AND E_NUMECHE='+IntToStr(TheTob.Detail[G.Row-1].GetValue('E_NUMECHE'));
  St:=St+' AND E_QUALIFPIECE="'+TheTob.Detail[G.Row-1].GetValue('E_QUALIFPIECE')+'" ' ;
  ExecuteSql(St) ;
  end ;
end ;

procedure TOF_VerifRib.OnImprimerClick(Sender : TObject);
begin
{$IFDEF EAGLCLIENT}
  PGIInfo('Non implémenté en mode CWAS', Ecran.Caption ) ;
{$ELSE}
  if gbIBAN then
    PrintDBGrid(G,Nil,'IBAN erronés','')
  else
    PrintDBGrid(G,Nil,'Relevés d''identités bancaires erronés','') ;
{$ENDIF}
End;

procedure TOF_VerifRib.OnGridDblClick(Sender : TObject);
begin MajGrid(False) ; end ;

procedure TOF_VerifRib.OnAffCorClick(Sender: TObject);
var i : integer ; BButton : TButton ;
begin
BButton:=TButton(GetControl('BAFFCOR')) ;
if CorAffOk then begin nb:=18 ; BButton.Hint:=TraduireMemoire('Enlever les '+gszRIB+' Corrects') ; end
            else begin nb:=0 ; BButton.Hint:=TraduireMemoire('Afficher les '+gszRIB+' Corrects') ; end ;
for i:=1 to G.RowCount-1 do if G.Cells[CL_ERREUR,i]=Msg[0] then G.RowHeights[i]:=nb ;
CorAffOk:=not CorAffOk ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 10/04/2002
Modifié le ... : 18/08/2004
Suite ........ : - LG - 18/08/2004 - Suppression de la fct debutdemois pour 
Suite ........ : l'appel de la saisie bor, ne fct pas avec les exercices 
Suite ........ : decalees
Mots clefs ... : 
*****************************************************************}
procedure TOF_VerifRib.OnZoomClick(Sender: TObject);
Var Q : TQuery ;
    M : RMvt ;
    Trouv : Boolean ;
    St,sMode : String ;
{$IFDEF EAGLCLIENT}
{$ELSE}
{$IFNDEF IMP}
    P  : RParFolio ;
{$ENDIF}
{$ENDIF}
begin
if TheTob=nil then Exit ;
sMode:=TheTob.Detail[G.Row-1].GetValue('E_MODESAISIE') ;
If ((sMode<>'') and (sMode<>'-'))  Then
  begin
{$IFDEF EAGLCLIENT}
  PGIInfo('Non implémenté en mode CWAS', Ecran.Caption ) ;
{$ELSE}
{$IFNDEF IMP}
  FillChar(P, Sizeof(P), #0) ;
  P.ParPeriode:=TheTob.Detail[G.Row-1].GetValue('E_DATECOMPTABLE') ;
  P.ParCodeJal:=TheTob.Detail[G.Row-1].GetValue('E_JOURNAL') ;
  P.ParNumFolio:=TheTob.Detail[G.Row-1].GetValue('E_NUMEROPIECE');
  P.ParNumLigne:=TheTob.Detail[G.Row-1].GetValue('E_NUMLIGNE') ;
  ChargeSaisieFolio(P, taConsult) ;
{$ENDIF}
{$ENDIF}
  end
  else
  begin
  St:='SELECT ' + SQLForIdent( fbGene ) + ' FROM ECRITURE WHERE ' ;
  St:=St+'E_JOURNAL="'+TheTob.Detail[G.Row-1].GetValue('E_JOURNAL')+'" ' ;
  St:=St+'AND E_EXERCICE="'+TheTob.Detail[G.Row-1].GetValue('E_EXERCICE')+'" '  ;
  St:=St+'AND E_DATECOMPTABLE="'+UsDateTime(TheTob.Detail[G.Row-1].GetValue('E_DATECOMPTABLE'))+'" ' ;
  St:=St+'AND E_NUMEROPIECE="'+IntToStr(TheTob.Detail[G.Row-1].GetValue('E_NUMEROPIECE'))+'" ' ;
  Q:=OpenSQL(St,True) ;
  Trouv:=Not Q.EOF ; if Trouv then M:=MvtToIdent(Q,fbGene,False) ;
  Ferme(Q) ;
  M.NumLigVisu:=TheTob.Detail[G.Row-1].GetValue('E_NUMLIGNE') ;
  if Trouv then LanceSaisie(Nil,taConsult,M) ;
  end ;
end;

procedure TOF_VerifRib.MAJGrid(Tous : Boolean);
var Rib,Aux,OldMsg,NewMsg : string ; Okok : boolean ;
begin
  Okok:=False ; Rib:=G.Cells[CL_RIB,G.Row]; Aux:=TheTob.Detail[G.Row-1].GetValue('E_AUXILIAIRE') ;
  {JP 28/05/07 : FQ 17460 : Gestion des tic / tid}
  if Aux = '' then Aux := TheTob.Detail[G.Row - 1].GetValue('E_GENERAL');

  if (VH^.PaysLocalisation<>CodeISOES) and (Copy(Rib, 1, 1) = '*') and not gbIBAN then begin
    PGIBox('La pièce '+TheTob.Detail[G.Row-1].GetString('E_NUMEROPIECE')+' contient un iban. La modification du rib ne sera pas effectuée.', Ecran.Caption);
    Exit;
  end; //XVI 24/02/2005
  if (not Tous) and (ModifLeRib(Rib,Aux)) then Okok:=True ;
  if (    Tous) and (RecupLeRib(Rib,Aux)) then Okok:=True ;
  if Okok then
    begin
    MAJLigneEcriture(Rib,G.Row-1) ;
    OldMsg:=G.Cells[CL_ERREUR,G.Row] ;
    if (gbIBAN) or ((VH^.PaysLocalisation=CodeISOES) and (RIBestIBAN(Rib))) then
      NewMsg:=MsgErr(VerificationIban(Rib))
    else
      NewMsg:=MsgErr(VerificationRib(Rib)) ;
    G.Cells[CL_ERREUR,G.Row]:=NewMsg ;
    if (NewMsg<>OldMsg) and (NewMsg=MsgErr(0)) then NBCorrect:=NbCorrect+1 ;
    if (NewMsg<>OldMsg) and (OldMsg=MsgErr(0)) then NBCorrect:=NbCorrect-1 ;
    AfficheNombre ;
    G.Cells[CL_RIB,G.Row]:=Rib;
    end ;
end ;

function TOF_VerifRib.RecupLeRib(var Rib, Aux : string ) : Boolean;
var Q      : TQuery ;
    Etab,Guichet,NumCompte,Cle,Dom : String ;
begin
   if VH^.PaysLocalisation=CodeISOES then
   Begin
     Rib:=GetRibPrincipal(Aux) ;
     result:=(Trim(RIB)<>'') ;
   End else
   Begin
     Result:=False ;
     Q:=OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="'+Aux+'" AND R_PRINCIPAL="X"',True) ;
     if not Q.EOF then
       begin
       Etab:=Q.FindField('R_ETABBQ').AsString ;
       Guichet:=Q.FindField('R_GUICHET').AsString ;
       NumCompte:=Q.FindField('R_NUMEROCOMPTE').AsString ;
       Cle:=Q.FindField('R_CLERIB').AsString ;
       Dom:=Q.FindField('R_DOMICILIATION').AsString ;
       Result:=True ;
       end
       else
         Rib:='' ;
     Ferme(Q) ;
     Rib:=EncodeRIB(Etab,Guichet,NumCompte,Cle,Dom) ;
   End ; //XVI 24/02/2005
end ;

procedure TOF_VerifRib.OnArgument(S: String);
var
  St : String;
begin
  inherited;
  St := S;
  gbIBAN := (ReadTokenSt(St) = 'IBAN');
  if gbIBAN then begin
    Ecran.caption := 'Vérification des IBAN';
    gszRIB := 'IBAN' ;
    Msg[0] := 'IBAN Correct';
    Msg[1] := 'IBAN inexistant';
    S := St;
  end
  else gszRIB := 'RIB'
end;

Initialization
registerclasses([TOF_VerifRib]);
end.

