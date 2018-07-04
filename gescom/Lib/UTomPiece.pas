unit UTomPiece;

interface

uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls, Controls, Classes, Forms, SysUtils, ComCtrls,
      HCtrls, HEnt1, HCrc, HMsgBox, UTOM, UTob, FactUtil, Menus,
{$IFDEF EAGLCLIENT}
      MaineAGL,eFichList,
{$ELSE}
      HDB,db,Fiche, FichList, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Hqry, Fe_Main, DBCtrls,
{$ENDIF}
{$IFDEF GPAO}
	   wCommuns,
{$ENDIF}
{$IFDEF BTP}
      FGestAffDet,
{$ENDIF}
      AglInit,Ent1, EntGC, LicUtil, graphics, M3FP,HRichOle,HTB97 ;

Type
     TOM_ParPiece = Class (TOM)
       private
           TobPiece : TOB;
           TOBDetailPiece : TOB;
{$IFDEF GPAO}
           OnlyPieceGP: Boolean;
{$ENDIF}
{$IFDEF BTP}
           BModifAffDet : TToolbarButton97;
           Souche       : THValComboBox;
           BValider     : TToolBarButton97;
{$ENDIF}

           procedure InitCombosAAucun ;
           function VerifNaturePiece (st_Champs, st_NatPiece : string;
                                      i_NumMess : integer) : Boolean;
           function ReferenceCirculaire (st_Champs, st_NatPiece : string) : Boolean;
           function Circularite(var st_NatPiece, Liste : string) : boolean;
           procedure AffecteValeurTableLibre(indice : integer);
           procedure ChangeTableLibre(indice : integer);
           procedure ClickOblTableLibre(indice : integer);
           function  SelectionNaturesParContextes(Plus : boolean = False) : string;
{$IFDEF BTP}
           procedure ChangeAffDefOuv(Sender: Tobject);
{$ENDIF}

           //AC 18/08/03 NV GESTION COMPTA DIFF
           procedure TstComptaStock;
           procedure ClickPassageCpta(Sender: Tobject);
           //Fin AC
{$IFDEF GPAO}
           procedure GPCacheChamps;
           procedure SetConfigState;

{$ENDIF}
           procedure ParamSaisie (Sender : Tobject);
       public
           procedure OnArgument (stArgument : string); override;
           procedure OnChangeField (F : TField)  ; override ;
           procedure OnUpdateRecord  ; override ;
           procedure OnLoadRecord  ; override ;
           procedure OnNewRecord  ; override ;
           procedure OnDeleteRecord  ; override ;
           procedure OnClose  ; override ;
     END ;

Type
     TOM_DomainePiece = Class (TOM)
       private
       public
           procedure OnChangeField (F : TField)  ; override ;
           procedure OnUpdateRecord  ; override ;
     END ;

const
	// libellés des messages
	TexteMessage: array[1..17] of string 	= (
          {1}        'Le code "Nature de Pièce" doit être renseigné',
          {2}        'L''intitulé doit être renseigné',
          {3}        'Le type de flux doit être renseigné' ,
          {4}        'Le mode de création doit être renseigné' ,
          {5}        'La nature suivante XXX n''est pas une nature de pièce existante',
          {6}        'Les pièces suivantes sont circulaires : XXX',
          {7}        'La nature équivalente XXX n''est pas une nature de pièce existante',
          {8}        'Cette nature de pièce ne peut être supprimée, des pièces y font références',
          {9}        'Suppression impossible, cette nature est référencée comme nature suivante pour le type de pièce XXX',
          {10}       'Le montant de visa doit être supérieur ou égal à zéro',
          {11}       'Vos modifications seront actives à la prochaine connexion',
          {12}       'Le domaine d''activité doit être renseigné',
          {13}       'Vous ne pouvez pas utiliser les tables libres si aucune d''elles n''est renseignée',
          {14}       'Un journal de type bordereau ne peut accepter des écritures de simulations',
          {15}       '',
//Message ParPiece
          {16}       'Vous-vous appliquer ce paramétrage à l''ensemble des natures de pièces concernées ?',
          {17}       'La mise à jour ne c''est pas faites !'

                     );

var
    EstModifie : boolean;

implementation

uses ParamSoc,
     SOUCHE_TOM
     ;

/////////////////////////////////////////////
// ****** TOM DOMAINE Piece ***********************
/////////////////////////////////////////////
procedure TOM_DomainePiece.OnChangeField(F: TField)  ;
begin
if (F.FieldName = 'GDP_MONTANTVISA') and (F.Value < 0.0) and (F.Value <> Null)then
    begin
    SetFocusControl('GDP_MONTANTVISA');
    LastError:=10 ; LastErrorMsg:=TexteMessage[LastError] ;
    exit ;
    end;
end;

procedure TOM_DomainePiece.OnUpdateRecord  ;
begin
    Inherited;
If (GetField ('GDP_DOMAINE')= '') then
    begin
    SetFocusControl ('GDP_DOMAINE');
    LastError:=12 ; LastErrorMsg:=TexteMessage[LastError] ;
    Exit;
    end;
If (GetField('GDP_LIBELLE') = '') then
    begin
    SetFocusControl('GDP_LIBELLE');
    LastError:=2 ; LastErrorMsg:=TexteMessage[LastError] ;
    Exit;
    end;
if (GetField('GDP_VISA')='X') and (GetField ('GDP_MONTANTVISA') < 0.0) then
    begin
    SetFocusControl('GDP_MONTANTVISA');
    LastError:=10 ; LastErrorMsg:=TexteMessage[LastError] ;
    Exit ;
    end;
end;


/////////////////////////////////////////////
// ****** TOM ParPiece ***********************
/////////////////////////////////////////////
procedure TOM_ParPiece.InitCombosAAucun ;
var CO_Control : TControl;
    Q_Combo : TQuery;
begin
    // recherche de tous les champs de type COMBO de la table GPP
    Q_Combo := OpenSql ('Select DH_NOMCHAMP from DECHAMPS ' +
                        'Where DH_PREFIXE = "GPP" AND ' +
                        'DH_TYPECHAMP = "COMBO"', FALSE);
    While not Q_Combo.Eof do
    begin
        //Q_Combo.Edit;
        CO_Control:=GetControl(Q_Combo.Fields[0].AsString) ;
        if (CO_Control is THDBValComboBox) then
        begin
            // Affectation par défaut du code AUCun ou RIEn
            if (THDBValComboBox(CO_Control).Values.Indexof('RIE')>-1) then
            begin
                THDBValComboBox(CO_Control).Value:='RIE'
            end
            else if (THDBValComboBox(CO_Control).Values.Indexof('AUC')>-1) then
            begin
                THDBValComboBox(CO_Control).Value:='AUC' ;
            end;
        end ;
        Q_Combo.Next;
    end;
    ferme (Q_Combo);
end ;

/////////////////////////////////////////////////////////////////////////////

function TOM_ParPiece.VerifNaturePiece (st_Champs, st_NatPiece : string;
                                        i_NumMess : integer) : Boolean;
var st_valeur : String ;
    i_ind : integer;
begin
    Result := True;
    st_valeur := ReadTokenSt (st_Champs) ;
    While st_valeur <> '' do
    begin
        if st_valeur <> st_NatPiece then
        begin
            if not (Presence ('PARPIECE', 'GPP_NATUREPIECEG', st_valeur)) then
            begin
                LastError := i_NumMess ;
                LastErrorMsg := TexteMessage[LastError];
                i_ind := pos ('XXX', LastErrorMsg);
                if i_ind > 0 then
                    LastErrorMsg := StringReplace (LastErrorMsg, 'XXX',
                                                   st_valeur, [rfIgnoreCase]);
                Result := False;
                exit;
            end;
        end;
        st_valeur := ReadTokenSt (st_Champs) ;
    end;
end;

/////////////////////////////////////////////////////////////////////////////
function TOM_ParPiece.ReferenceCirculaire (st_Champs, st_NatPiece : string) : Boolean;
var Liste, RefCirc : string ;
    i_ind : integer ;
    Q_Nature : TQuery;
begin
Result:=True ;
Q_Nature:=OpenSQL('SELECT GPP_NATUREPIECEG, GPP_NATURESUIVANTE FROM PARPIECE WHERE GPP_NATUREPIECEG<>"'+st_NatPiece+'"',True) ;
TobPiece:=TOB.Create('PARPIECE',Nil,-1);
TobPiece.LoadDetailDB('PARPIECE','','',Q_Nature,False);
Ferme(Q_Nature);
Tob.Create('PARPIECE',TobPiece,-1);
With Tob.Create('PARPIECE',TobPiece,-1) do
     begin
     AddChampSup('GPP_NATUREPIECEG',false);
     AddChampSup('GPP_NATURESUIVANTE',false);
     PutValue('GPP_NATUREPIECEG',st_NatPiece);
     PutValue('GPP_NATURESUIVANTE',st_Champs);
     end;
if Circularite(st_NatPiece, Liste) then
    begin
    While Liste<>'' do
        begin
        RefCirc:=RefCirc+ReadTokenSt(Liste) ;
        if Liste<>'' then RefCirc:=RefCirc+' -> ';
        end ;
    LastError := 6 ;
    LastErrorMsg := TexteMessage[LastError];
    i_ind := pos ('XXX', LastErrorMsg);
    if i_ind > 0 then
            LastErrorMsg := StringReplace (LastErrorMsg,'XXX',RefCirc, [rfIgnoreCase]);
    Result := False;
    end;
TobPiece.free;
end ;


function TOM_ParPiece.Circularite(var st_NatPiece, Liste : string) : boolean;
var St, NatSuiv : string;
   TobPieceSuiv : TOB ;
begin
Result:=False ; Exit ;
// JLD Pour contourner bug
Liste:=Liste+st_NatPiece+';';
Result:=False;
TobPieceSuiv:=TobPiece.FindFirst(['GPP_NATUREPIECEG'],[st_NatPiece],false);
if  TobPieceSuiv<>Nil then NatSuiv:=TobPieceSuiv.GetValue('GPP_NATURESUIVANTE')
else NatSuiv:='';
While NatSuiv<>''  do
      begin
      St:=ReadTokenSt(NatSuiv);
      if st_NatPiece=St then continue;
      if pos(St,Liste)=0 then Result:=false
         else begin Result:=true ; Liste:=Liste+St ; exit ; end ;
      if Circularite(St, Liste)=True then begin Result:=True ; exit ; end ;
      end;
TobPieceSuiv.free;
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_ParPiece.AffecteValeurTableLibre(indice : integer);
var  St_Plus, St_Value, St : string;
     i_ind : integer;
begin
if (ctxMode  in V_PGI.PGIContexte) then
   BEGIN
    St_Value := THDBValComboBox(GetControl('GPP_PIECETABLE'+InttoStr(Indice))).Value;
    For i_ind := 1 to 3 do
        BEGIN
        if i_ind = Indice then continue;
        St := THDBValComboBox(GetControl('GPP_PIECETABLE'+InttoStr(i_ind))).Value;
        If St <> '' then St_Plus := St_Plus + ' AND CC_CODE <>"' + St + '"';
        END;
    SetControlProperty('GPP_PIECETABLE'+InttoStr(Indice),'Plus',St_Plus);
    SetControlProperty('GPP_PIECETABLE'+InttoStr(Indice),'Value',St_Value);
   end;
end;

procedure TOM_ParPiece.ChangeTableLibre(indice : integer);
var  St_Plus : string;
     i_ind : integer;
begin
if (ctxMode  in V_PGI.PGIContexte) then
   BEGIN
    St_Plus := THDBValComboBox(GetControl('GPP_PIECETABLE'+InttoStr(Indice))).Value;
    if St_Plus = '' then
        BEGIN
        SetControlEnabled('GPP_CODPIECEDEF'+InttoStr(Indice),True); SetControlProperty('GPP_CODPIECEDEF'+InttoStr(Indice),'Value','');
        SetControlEnabled('GPP_CODPIECEDEF'+InttoStr(Indice),False); SetControlProperty('GPP_CODPIECEDEF'+InttoStr(Indice),'Color',clBtnFace);
        SetControlEnabled('GPP_CODEPIECEOBL'+InttoStr(Indice),True); SetControlChecked('GPP_CODEPIECEOBL'+InttoStr(Indice),False);
        SetControlEnabled('GPP_CODEPIECEOBL'+InttoStr(Indice),False);
        For i_ind := indice + 1 to 3 do
            BEGIN
            SetControlEnabled('GPP_PIECETABLE'+InttoStr(i_ind),True); SetControlProperty('GPP_PIECETABLE'+InttoStr(i_ind),'Value','');
            SetControlEnabled('GPP_PIECETABLE'+InttoStr(i_ind),False); SetControlProperty('GPP_PIECETABLE'+InttoStr(i_ind),'Color',clBtnFace);
            SetControlEnabled('GPP_CODPIECEDEF'+InttoStr(i_ind),True); SetControlProperty('GPP_CODPIECEDEF'+InttoStr(i_ind),'Value','');
            SetControlEnabled('GPP_CODPIECEDEF'+InttoStr(i_ind),False); SetControlProperty('GPP_CODPIECEDEF'+InttoStr(i_ind),'Color',clBtnFace);
            SetControlEnabled('GPP_CODEPIECEOBL'+InttoStr(i_ind),True); SetControlChecked('GPP_CODEPIECEOBL'+InttoStr(i_ind),False);
            SetControlEnabled('GPP_CODEPIECEOBL'+InttoStr(i_ind),False);
            END;
        END else
        BEGIN
        SetControlEnabled('GPP_CODPIECEDEF'+InttoStr(Indice),True); SetControlProperty('GPP_CODPIECEDEF'+InttoStr(Indice),'Color',clwindow);
        SetControlProperty('GPP_CODPIECEDEF'+InttoStr(Indice),'Plus',St_Plus);
        SetControlEnabled('GPP_CODEPIECEOBL'+InttoStr(Indice),True); SetControlChecked('GPP_CODEPIECEOBL'+InttoStr(indice),False);
        if Indice < 3 then
            BEGIN
            SetControlEnabled('GPP_PIECETABLE'+InttoStr(Indice+1),True); SetControlProperty('GPP_PIECETABLE'+InttoStr(Indice+1),'Color',clwindow);
            END;
        END;
    end;
// Modif du 01/08/2001
{
if THDBValComboBox(GetControl('GPP_PIECETABLE1')).Value='' then
  begin
  SetControlEnabled('GPP_AFFPIECETABLE',True); SetControlChecked('GPP_AFFPIECETABLE',False);
  SetControlEnabled('GPP_AFFPIECETABLE',False);
  end
else SetControlEnabled('GPP_AFFPIECETABLE',true);
}
end;

procedure TOM_ParPiece.ClickOblTableLibre(indice : integer);
Var NbrZone : integer;
    Oblig : Boolean;
begin
if (ctxMode  in V_PGI.PGIContexte) then
   BEGIN
{$IFDEF EAGLCLIENT}
    NbrZone:=THValComboBox(GetControl('GPP_CODPIECEDEF'+InttoStr(indice))).Items.Count;
    Oblig:=TCheckBox(GetControl('GPP_CODEPIECEOBL'+InttoStr(indice))).Checked;
{$ELSE}
    NbrZone:=THDBValComboBox(GetControl('GPP_CODPIECEDEF'+InttoStr(indice))).Items.Count;
    Oblig:=TDBCheckBox(GetControl('GPP_CODEPIECEOBL'+InttoStr(indice))).Checked;
{$ENDIF}
    if NbrZone=1 then
      begin
      if Oblig then
        begin
        PGIBox('La table libre doit posséder au moins une valeur pour qu''elle puisse être rendue obligatoire',Ecran.Caption);
        SetField('GPP_CODEPIECEOBL'+InttoStr(indice),'-');
        //SetControlChecked('GPP_CODEPIECEOBL'+InttoStr(indice),False);
        SetFocusControl('GPP_CODEPIECEOBL'+InttoStr(indice));
        end;
      end;
   end;
end;

function TOM_ParPiece.SelectionNaturesParContextes(Plus : boolean = False) : string;
Var ind1 : integer;
    stUp, stArg : string;
    TobTemp : TOB;

    function TesteContexte(st1 : string) : boolean;
    begin
    Result := False;
    if st1 = '' then Result := True;
    if Pos('GC', st1)  <> 0 then Result := True;
    if Pos('NEG', st1) <> 0 then Result := True;
    if Pos('AFF', st1) <> 0 then Result := True;
    if Pos('SCO', st1) <> 0 then Result := True;
    if Pos('TEM', st1) <> 0 then Result := True;
    if Pos('GCA', st1) <> 0 then Result := True;
    if Pos('GRC', st1) <> 0 then Result := True;
{$IFDEF GPAO} 
    if OnlyPieceGP then						{ Ne montre que les pièces GP }
       Result := Pos('GPAO', st1) <> 0
    else											{ Montre toutes les pièces sauf celles pilotées par la GP }
       if Pos('GPAO', st1) <> 0 then Result := False;
{$ENDIF}
    end;

begin
//  Valeurs à saisir pour les contextes par defaut
//  GC  = ctxGescom,
//  NEG = ctxNegoce,
//  AFF = ctxAffaire,
//  SCO = ctxScot,
//  TEM = ctxTempo,
//  MOD = ctxMode,
//  GCA = ctxGCAff,
//  GRC = ctxGRC,
//  CHR = ctxCHR,
//  BTP = ctxBTP,
//
//  Plus = Vrai      : appel pour tablettes
Result := '';
if ctxMode in V_PGI.PGIContexte then Exit;
if ctxBTP in V_PGI.PGIContexte then Exit;
if ctxCHR in V_PGI.PGIContexte then Exit;
//  On a Plus a Vrai, donc on genere une selection à appliquer sur une tablette
if Plus then
    begin
    Result := 'GPP_CONTEXTES="" ';
    stUp := 'or GPP_CONTEXTES like "%';
    stArg := '%" ';
    if ctxGescom  in V_PGI.PGIContexte then Result := Result + stUp + 'GC' + stArg;
    if ctxNegoce  in V_PGI.PGIContexte then Result := Result + stUp + 'NEG' + stArg;
    if ctxAffaire in V_PGI.PGIContexte then Result := Result + stUp + 'AFF' + stArg;
    if ctxScot    in V_PGI.PGIContexte then Result := Result + stUp + 'SCO' + stArg;
    if ctxTempo   in V_PGI.PGIContexte then Result := Result + stUp + 'TEM' + stArg;
    if ctxGCAff   in V_PGI.PGIContexte then Result := Result + stUp + 'GCA' + stArg;
    if ctxGRC     in V_PGI.PGIContexte then Result := Result + stUp + 'GRC' + stArg;
{$IFDEF GPAO}
    if ctxGPAO    in V_PGI.PGIContexte then Result := Result + stUp + 'GP' + stArg;
{$ENDIF}
    if ctxBTP    in V_PGI.PGIContexte then Result := Result + stUp + 'BTP' + stArg;
    Exit;
    end;
//  On vient de GCPARPIECE, on sort avec Result vide pour appliquer aucune selection
if Ecran.Name = 'GCPARPIECE' then Exit;
//  sinon on genere une chaine compatible eagl ou pas en fonction...
for ind1 := 0 to VH_GC.TOBParPiece.Detail.Count - 1 do
    begin
    TobTemp := VH_GC.TOBParPiece.Detail[ind1];
    if not TesteContexte(TobTemp.GetValue('GPP_CONTEXTES')) then Continue;
    if Result = '' then
{$IFDEF EAGLCLIENT}
        Result := TobTemp.GetValue('GPP_NATUREPIECEG')
{$ELSE}
        Result := 'GPP_NATUREPIECEG = ''' + TobTemp.GetValue('GPP_NATUREPIECEG') + ''''
{$ENDIF}
        else
{$IFDEF EAGLCLIENT}
        Result := Result + '" or GPP_NATUREPIECEG = "' + TobTemp.GetValue('GPP_NATUREPIECEG');
{$ELSE}
        Result := Result + ' or GPP_NATUREPIECEG = ''' + TobTemp.GetValue('GPP_NATUREPIECEG') + '''';
{$ENDIF}
    end;
{$IFDEF EAGLCLIENT}
Result := Result + ';';
{$ENDIF}
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_ParPiece.OnChangeField(F: TField)  ;
begin
if (F.FieldName = 'GPP_MONTANTVISA') and (F.Value < 0.0) and (F.Value <> Null)then
    begin
    SetFocusControl('GPP_MONTANTVISA');
    LastError:=10 ;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit ;
    end else //AC 18/08/03 NV GESTION COMPTA DIFF
    if F.FieldName = 'GPP_TYPEECRCPTA' then
    begin
    	TstComptaStock;
    end else if (F.FieldName = 'GPP_VENTEACHAT') then
    begin
    	if (GetField('GPP_VENTEACHAT')='ACH') then
      begin
      	TDBCheckBox(GetControl('GPP_FORCEENPA')).Checked:=false;
        TDBCheckBox(GetControl('GPP_FORCEENPA')).Visible := false;
      end else
      begin
        TDBCheckBox(GetControl('GPP_FORCEENPA')).Visible := True;
      end;
    end;

    // Fin AC
end;

/////////////////////////////////////////////////////////////////////////////

procedure TOM_ParPiece.OnUpdateRecord  ;
var st_NatPiece : string;
    st_NatSaisie: string;
    st_message  : String ;
    i_error     : integer;
    QQ          : Tquery;
    ModeSaisie  : string;
begin
    Inherited;
If (GetField ('GPP_NATUREPIECEG')= '') then
    begin
    SetFocusControl ('GPP_NATUREPIECEG');
    LastError := 1 ;
    LastErrorMsg := TexteMessage[LastError] ;
    exit;
    end;

If (GetField ('GPP_LIBELLE') = '') then
    begin
    SetFocusControl ('GPP_LIBELLE');
    LastError := 2 ;
    LastErrorMsg := TexteMessage[LastError] ;
    exit;
    end;
If (GetField ('GCVENTEACHAT') = '') then
    begin
    SetFocusControl ('GCVENTEACHAT');
    LastError := 3 ;
    LastErrorMsg := TexteMessage[LastError] ;
    exit;
    end;
if (GetField('GPP_VISA') = 'X') and (GetField ('GPP_MONTANTVISA') < 0.0) then
    begin
    SetFocusControl('GPP_MONTANTVISA');
    LastError:=10 ;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit ;
    end;
st_NatPiece := GetField ('GPP_NATUREPIECEG');
st_NatSaisie := GetField ('GPP_NATURESUIVANTE');
if (Pos (st_NatPiece, st_NatSaisie) = 0) then
    begin
    st_NatSaisie := st_NatPiece + ';' + st_NatSaisie;
    SetField ('GPP_NATURESUIVANTE', st_NatSaisie)
    end;
    // Vérification de l'existance des natures de pièces suivantes
If not (VerifNaturePiece (st_NatSaisie, st_NatPiece, 5)) then
    begin
    i_error := LastError ;
    st_message := LastErrorMsg;
    SetFocusControl ('GPP_NATURESUIVANTE');
    LastError := i_error ;
    LastErrorMsg := st_message;
    exit;
    end;
// recherche des référencements circulaires
If not (ReferenceCirculaire (st_NatSaisie, st_NatPiece)) then
    begin
    i_error := LastError ;
    st_message := LastErrorMsg;
    SetFocusControl ('GPP_NATURESUIVANTE');
    LastError := i_error ;
    LastErrorMsg := st_message;
    exit;
    end;
// Vérification de l'existance des natures de pièces équivalentes
st_NatSaisie := GetField ('GPP_EQUIPIECE');
If not (VerifNaturePiece (st_NatSaisie, st_NatPiece, 7)) then
    begin
    i_error := LastError ;
    st_message := LastErrorMsg;
    SetFocusControl ('GPP_EQUIPIECE');
    LastError := i_error ;
    LastErrorMsg := st_message;
    exit;
    end;
{$IFNDEF EAGLCLIENT}
  SetField ('GPP_CRC', Integer (GetCRC32ForData (DS))) ;
{$ENDIF}
// Modif du 01/08/2001
if (TCheckBox(GetControl('GPP_AFFPIECETABLE')).Checked) and (GetField('GPP_PIECETABLE1')='') then
    begin
    SetFocusControl('GPP_AFFPIECETABLE');
    LastError:=13 ;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit ;
    end;
//AC 18/08/03 NV GESTION COMPTA DIFF
if GetField('GPP_TYPEECRCPTA') = '' then
    SetField('GPP_TYPEECRCPTA', 'RIE');
if GetField('GPP_TYPEPASSCPTA') = '' then
    SetField('GPP_TYPEPASSCPTA', 'REE');
if GetField('GPP_TYPEPASSACC') = '' then
    SetField('GPP_TYPEPASSACC', 'REE');

if GetField('GPP_TYPEECRCPTA') = 'SIM' then
begin
	// on verifie que le journal n'est pas en mode bordereau
  QQ := OpenSql('Select J_MODESAISIE from JOURNAL where J_JOURNAL="' + GetField('GPP_JOURNALCPTA') + '"', TRUE);
  if not QQ.EOF then
  begin
    ModeSaisie := QQ.findField('J_MODESAISIE').asString;
    ferme (QQ);
    if ModeSaisie <> '-' then
    begin
      SetFocusControl ('GPP_EQUIPIECE');
      LastError:=14 ;
      LastErrorMsg:=TexteMessage[LastError] ;
      exit;
    end;
  end else Ferme (QQ);

end;

//Fin AC
EstModifie := True;
end;

//////////////////////////////////////////////////////////////////////

procedure TOM_ParPiece.OnLoadRecord  ;
var i_ind   : Integer;
    i       : integer ;
    st_Flux : String;
    st_Plus : string;
    StSQL   : String;
    Year,Month,Day : Word;
begin
Inherited;
TOBDetailPiece.ClearDetail;
//
st_Flux := GetField ('GPP_VENTEACHAT');
if st_Flux='VEN' then
   BEGIN
   SetControlProperty('GPP_JOURNALCPTA','Plus','AND J_NATUREJAL="VTE"');
   if Ecran.Name = 'GCPARPIECEUSR' then
      BEGIN
      SetControlProperty('GPP_JOURNALCPTA','Enabled',True);
      SetControlProperty('GPP_JOURNALCPTA','Color',clwindow);
      END;
   END else
   if st_Flux='ACH' then
      BEGIN
      SetControlProperty('GPP_FORCEENPA','visible',false);
      SetControlProperty('GPP_JOURNALCPTA','Plus','AND J_NATUREJAL="ACH"');
      if Ecran.Name = 'GCPARPIECEUSR' then
         BEGIN
         SetControlProperty('GPP_JOURNALCPTA','Enabled',True);
         SetControlProperty('GPP_JOURNALCPTA','Color',clwindow);
         END;
      END else
      BEGIN
      SetControlProperty('GPP_JOURNALCPTA','Plus','');
      if Ecran.Name = 'GCPARPIECEUSR' then
            BEGIN
            SetControlProperty('GPP_JOURNALCPTA','Enabled',False);
            SetControlProperty('GPP_JOURNALCPTA','Color',clbtnface);
            END;
      END;
if (not (ctxAffaire in V_PGI.PGIContexte))and (Ecran.Name='GCPARPIECEUSR') then
   if (GetField ('GPP_NATUREPIECEG')='CC') or (GetField ('GPP_NATUREPIECEG')='DE') then
      SetControlVisible ('PCONTREMARQUE', True)
      else
      SetControlVisible ('PCONTREMARQUE', False);

i:=0;
for i_ind:=1 to 8 do
    begin
    SetControlProperty('GPP_IFL'+InttoStr(i_ind),'Plus','AND CO_CODE NOT LIKE "9%"');
    if GetField('GPP_IFL'+InttoStr(i_ind))<>'' then
       begin
       SetControlProperty('GPP_IFL'+InttoStr(i_ind),'Color',clwindow);
       SetControlEnabled('GPP_IFL'+InttoStr(i_ind),true);
       i:=i_ind;
       end else
       begin
       SetControlProperty('GPP_IFL'+InttoStr(i_ind),'Color',clbtnface);
       SetControlEnabled('GPP_IFL'+InttoStr(i_ind),false);
       end;
    end;
if i<8 then
   begin
   SetControlProperty('GPP_IFL'+InttoStr(i+1),'Color',clwindow);
   SetControlEnabled('GPP_IFL'+InttoStr(i+1),true);
   end;
//Gestion : Tables libres
if (ctxMode  in V_PGI.PGIContexte) then
   BEGIN
 // Modif 01/08/2001
 //  if GetField('GPP_PIECETABLE1')='' then SetControlEnabled('GPP_AFFPIECETABLE',False)
 //  else SetControlEnabled('GPP_AFFPIECETABLE',true);
   i:=0;
   for i_ind:=1 to 3 do
       begin
       St_Plus := THDBValComboBox(GetControl('GPP_PIECETABLE'+InttoStr(i_ind))).Value;
       //if GetField('GPP_PIECETABLE'+InttoStr(i_ind))<>'' then
       if St_Plus<>'' then
          begin
          SetControlEnabled('GPP_PIECETABLE'+InttoStr(i_ind),true); SetControlProperty('GPP_PIECETABLE'+InttoStr(i_ind),'Color',clwindow);
          SetControlEnabled('GPP_CODPIECEDEF'+InttoStr(i_ind),true); SetControlProperty('GPP_CODPIECEDEF'+InttoStr(i_ind),'Color',clwindow);
          SetControlProperty('GPP_CODPIECEDEF'+InttoStr(i_ind),'Plus',St_Plus);
          SetControlEnabled('GPP_CODEPIECEOBL'+InttoStr(i_ind),true);
          i:=i_ind;
          end else
          begin
          SetControlEnabled('GPP_PIECETABLE'+InttoStr(i_ind),false); SetControlProperty('GPP_PIECETABLE'+InttoStr(i_ind),'Color',clbtnface);
          SetControlEnabled('GPP_CODPIECEDEF'+InttoStr(i_ind),false); SetControlProperty('GPP_CODPIECEDEF'+InttoStr(i_ind),'Color',clbtnface);
          SetControlEnabled('GPP_CODEPIECEOBL'+InttoStr(i_ind),false);
          end;
       end;
   if i<3 then
      begin
      SetControlEnabled('GPP_PIECETABLE'+InttoStr(i+1),true); SetControlProperty('GPP_PIECETABLE'+InttoStr(i+1),'Color',clwindow);
      end;
   END ;
{$IFDEF BTP}
   THRichEditOle(GetControl('TMODEAFFDEF')).lines.Text  := RecupModeAffLibelles (GetField('GPP_TYPEPRESENT'));
{$ENDIF}
{$IFDEF GPAO}
  SetConfigState;
{$ENDIF}
//AC 18/08/03 NV GESTION COMPTA DIFF
   TstComptaStock ;
//Fin AC
  {$IFDEF CCS3}
  if (ctxAffaire  in V_PGI.PGIContexte) then   SetControlVisible ('PAFFAIRE', True);
  {$ENDIF}
end;


procedure TOM_ParPiece.OnNewRecord  ;
begin
    InHerited ;
    SetActiveTabSheet ('PGENERALE');
    InitCombosAAucun ;
    SetControlChecked ('GPP_HISTORIQUE', TRUE);
    SetField ('GPP_SENSPIECE', 'MIX');
    SetField ('GPP_IMPIMMEDIATE', 'X');
    SetField ('GPP_IMPMODELE',VH_GC.GCImpModeleDefaut );
{$IFDEF BTP}
    SetField('GPP_TYPEPRESENT',DOU_AUCUN);
    THRichEditOle(GetControl('TMODEAFFDEF')).lines.text :=RecupModeAffLibelles (GetField('GPP_TYPEPRESENT'));
{$ENDIF}
    SetField('GPP_PIECEPILOTE', '-');
    SetField('GPP_CFGART', '-'); 
{$IFDEF GPAO} 
    if OnlyPieceGP then SetField('GPP_PIECEPILOTE', 'X')
{$ENDIF}
end;

procedure TOM_ParPiece.OnDeleteRecord  ;
var Q_Piece : TQuery;
    Trouve : Boolean;
    st_NatPiece, st_Ref : string;
    i_ind : integer;
begin
    // Modif le 01/08/2001
    st_NatPiece := GetField ('GPP_NATUREPIECEG');
    Trouve:=ExisteSQL('SELECT GP_NATUREPIECEG FROM PIECE ' +
                     'WHERE GP_NATUREPIECEG="' + st_NatPiece + '"') ;
    if Trouve then
    begin
        LastError := 8 ;
        LastErrorMsg := TexteMessage[LastError] ;
        exit;
    end;
    Q_Piece:=OpenSQL('SELECT GPP_NATUREPIECEG FROM PARPIECE ' +
                     'WHERE GPP_NATURESUIVANTE LIKE "%' + st_NatPiece + '%" AND ' +
                     'GPP_NATUREPIECEG <> "' + st_NatPiece + '"', True) ;
    if (not Q_Piece.Eof) then
    begin
        st_ref := Q_Piece.Fields[0].AsString;
        LastError := 9 ;
        LastErrorMsg := TexteMessage[LastError] ;
        i_ind := pos ('XXX', LastErrorMsg);
        if i_ind > 0 then
            LastErrorMsg := StringReplace (LastErrorMsg, 'XXX',
                                           st_Ref, [rfIgnoreCase]);
        Ferme (Q_Piece);
        exit;
    end;
    Ferme (Q_Piece);
end;

procedure TOM_ParPiece.OnArgument (stArgument : string);
VAR F : TForm ;
    stRange : string;
    Pop : TPopupMenu;
    Cpt : integer;
{$IFDEF NOMADE}
    bMenuSiteOk : boolean ;
{$ENDIF}
BEGIN
  TOBDetailPiece := TOB.Create ('LES PIECES',nil,-1);
  SetControlVisible('MnBSV',VH_GC.ISGEDBSV);
  {$IFDEF GPAO}
  { Appel du ParPiece pour paramètrage des pièces GP }
  OnlyPieceGP := GetArgumentValue(stArgument, 'PARAMPIECESGP') = 'OUI';
  {$ENDIF}
  F := TForm(Ecran);
  stRange := SelectionNaturesParContextes(False);
  if stRange <> '' then
{$IFDEF EAGLCLIENT}
    TFFicheListe(Ecran).FRange := stRange;
{$ELSE}
  begin
    Ds.Filter := stRange;
    DS.Filtered := True;
  end;
{$ENDIF}
  {$IFDEF MODE}
  // Pour la mode il faut le bon modele d'étiquette trouvé par la tablette GCIMPETIQMODE
  if (ctxMode in V_PGI.PGIContexte) then
  THValComboBox(F.FindComponent ('GPP_ETATETIQ')).DataType:='GCIMPETIQMODE';
  SetControlVisible ('GPP_MAJINFOTIERS', True);
  {$ELSE}
  //  SetControlVisible ('GPP_MAJINFOTIERS', False);
  SetControlVisible ('GPP_MAJINFOTIERS', GetParamSoc('SO_GCTRV'));
  {$ENDIF}
  // suppression de certains onglets en gestion d'affaires
  if ctxAffaire in V_PGI.PGIContexte then
  BEGIN
    SetControlVisible ('PSTOCK', False);  SetControlVisible ('PCONTREMARQUE', False);
    If CtxScot in V_Pgi.PgiContexte then SetControlVisible ('PPREF', False); //mcd 25/09 vu en GA, pas en GI
    SetControlVisible ('PPORT', False);
    {$IFNDEF BTP}
    SetControlVisible ('GPP_LISTESAISIE',False);
    SetControlVisible ('GPP_LISTEAFFAIRE',True);
    {$ELSE}
    SetControlVisible ('GPP_LISTESAISIE',true);
    SetControlVisible ('TGPP_LISTESAISIE',true);
    SetControlVisible ('GPP_LISTEAFFAIRE',false);
    SetControlVisible ('PSTOCK', True);  SetControlVisible ('PCONTREMARQUE', False);
    SetControlVisible ('PPREF', True);
    {$IFDEF EAGLCLIENT}
    THEdit(GetCOntrol('GPP_LISTESAISIE')).DataType := 'BTLISTESAISIE';
    {$ELSE}
    THDBEdit(GetCOntrol('GPP_LISTESAISIE')).DataType := 'BTLISTESAISIE';
    {$ENDIF}
    {$ENDIF}
  END
  else
  BEGIN
    SetControlVisible ('GPP_LISTESAISIE',True);
    SetControlVisible ('TGPP_LISTESAISIE',True);
    SetControlVisible ('GPP_LISTEAFFAIRE',False);
    if not(ctxGCAFF in V_PGI.PGIContexte) then SetControlVisible ('PAFFAIRE',False);
    if not(ctxMode  in V_PGI.PGIContexte) then SetControlVisible ('GPIECETABLE',False);
    //  BBI, correction fiche 10336
    if not(ctxMode  in V_PGI.PGIContexte) then
    begin
    SetControlProperty ('GPP_AFFPIECETABLE', 'Caption', 'Tables libres pièces obligatoires');
    SetControlProperty ('GPP_AFFPIECETABLE', 'Width', 170);
    end;
    //  BBI, fin correction fiche 10336
  END;
  {$IFDEF CCS3}
  SetControlVisible ('GPP_LOT', False);
  SetControlVisible ('GPP_NUMEROSERIE', False);
  SetControlVisible('GPP_TYPEECRSTOCK',False) ; SetControlVisible('TGPP_TYPEECRSTOCK',False) ;
  {$ENDIF}
  if Ecran.Name='GCPARPIECEUSR' then
  begin
    {$IFDEF CCS3}
    SetControlVisible ('GPP_ECLATEAFFAIRE', False);
    {$ENDIF}
    SetControlVisible ('PCONTREMARQUE', False);
    SetControlEnabled('GPP_ESTAVOIR',False);
    if V_PGI.Superviseur then
    begin
      SetControlEnabled('GPP_TYPEPASSCPTA',True);
      SetControlProperty('GPP_TYPEPASSCPTA','Color',ClWindow);
    end;
    if (ctxMode in V_PGI.PGIContexte) or (ctxBTP in V_PGI.PGIContexte) or (ctxCHR in V_PGI.PGIContexte) then
    SetControlVisible ('GPP_ARTFOURPRIN', True) else SetControlVisible ('GPP_ARTFOURPRIN', False);
  end;
  {$IFDEF BTP}
  BModifAffDet := TToolBarButton97(Getcontrol('BMODIFMODEAFF'));
  BModifAffDet.OnClick := ChangeAffDefOuv;
  SetControlProperty ('PNOMENC','Caption','Ouvrages');
  SetControlProperty ('PPORTS','Caption','Ports, Frais et Retenues de Garantie');
  SetControlVisible ('GRETENU', True);
  SetControlVisible ('GPP_VISA', True);
  SetControlVisible ('GPP_MONTANTVISA', True);
  SetControlVisible ('TGPP_MONTANTVISA', True);
  {$ELSE}
  SetControlVisible ('PNOMENC', False);
  {$ENDIF}
  SetControlVisible ('TGPP_TYPEPRESDOC', True);
  SetControlVisible ('GPP_TYPEPRESDOC', True);

  {$IFDEF NOMADE}
  //Montre/cache ligne des exceptions par sites pour les documents
  if ( PCP_LesSites <> nil ) and ( PCP_LesSites.LeSiteLocal <> nil ) then
  bMenuSiteOk := ( PCP_LesSites.LeSiteLocal.LesVariables.Find('$REP') <> nil )
  else bMenuSiteOk := False ;

  if not bMenuSiteOk then
  begin
    pop := TPopupMenu( GetControl( 'POP_M' ) ) ;
    if pop <> nil then for Cpt := 0 to pop.items.count - 1 do
    if pop.items[Cpt].name = 'mnsite' then begin pop.items[Cpt].visible := False ; break ; end ;
  end
  else //Montre case à cocher transfert Vente ou Achat
  begin
    SetControlProperty( 'GPP_TRSFVENTE', 'Visible', True ) ;
    SetControlProperty( 'GPP_TRSFACHAT', 'Visible', True ) ; //Montre Achat mais désactive
    SetControlProperty( 'GPP_TRSFACHAT', 'Enabled', False ) ;
  end ;
  {$ELSE}
  //Cache ligne des exceptions par sites pour les documents
  pop := TPopupMenu( GetControl( 'POP_M' ) ) ;
  if pop <> nil then for Cpt := 0 to pop.items.count - 1 do
  if pop.items[Cpt].name = 'mnsite' then begin pop.items[Cpt].visible := False ; break ; end ;
  if CtxScot in V_PGI.PgiContexte then
  begin   // mcd 22/10/03 on cache domaine et site non géré en GI
    if pop <> nil then
    for Cpt := 0 to pop.items.count - 1 do
    begin
      if uppercase(pop.items[Cpt].name) = 'MNSITE' then  pop.items[Cpt].visible := False ;
      if uppercase(pop.items[Cpt].name) = 'MNDOMAINE' then  pop.items[Cpt].visible := False ;
    end;
  end;
  {$ENDIF}

  {$IFDEF GPAO}
  if OnlyPieceGP then
  begin
    { Masque les champs qui ne servent pas dans le paramètrage des pièces de la GP }
    GPCacheChamps;
    { Cache le bouton d'accès aux complément de paramètrage }
    if GetControl('BCOMPLEMENT') <> nil then SetControlVisible('BCOMPLEMENT', False);
  end;
  {$ENDIF}
  //AC 18/08/03 NV GESTION COMPTA DIFF
  if GetParamSoc('SO_GCDESACTIVECOMPTA') and (CtxMode in V_PGI.PGIContexte) then
  begin
    SetControlProperty('GPP_TYPEECRCPTA', 'Visible', False);
    SetControlProperty('GPP_TYPEPASSCPTA', 'Visible', False);
    SetControlProperty('GPP_TYPEPASSACC', 'Visible', False);
    SetControlProperty('GPP_TYPEECRSTOCK', 'Visible', False);
  end ;
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('GPP_TYPEECRCPTA')).OnClick := ClickPassageCpta;
  {$ELSE}
  THDBValComboBox(GetControl('GPP_TYPEECRCPTA')).OnClick := ClickPassageCpta;
  {$ENDIF}
  //Fin AC

  SetActiveTabSheet('PGENERALE');
  SetControlProperty('GBTARIFS','Visible',GetPAramSoc('SO_PREFSYSTTARIF'));

  {$IFDEF BTP}
  {$IFNDEF EAGLCLIENT}
  THDBValComboBox (GetCOntrol('GPP_TYPECOMMERCIAL')).Vide := true;
  THDBValComboBox (GetCOntrol('GPP_TYPECOMMERCIAL')).ReLoad;
  {$ELSE}
  THValComboBox (GetCOntrol('GPP_TYPECOMMERCIAL')).Vide := true;
  THValComboBox (GetCOntrol('GPP_TYPECOMMERCIAL')).ReLoad;
  {$ENDIF}
  // SetControlProperty ('GPP_TYPECOMMERCIAL','Vide',True);
  {$ENDIF}

  Souche        := THValComboBox(GetControl('GPP_SOUCHE'));
  BValider      := TToolBarButton97(GetControl('BVALIDER'));

  if TToolbarButton97 (GetControl('BPARAMSAISIE')) <> nil then
  begin
  {$IFDEF V10}                                               
    TToolbarButton97 (GetControl('BPARAMSAISIE')).OnClick := ParamSaisie;
  {$ELSE}
    TToolbarButton97 (GetControl('BPARAMSAISIE')).visible := false;
  {$ENDIF}
  end;

  {$IFNDEF V10}
  if TCheckBox (GetCOntrol('GPP_TAILLAGEAUTO')) <> nil then
  begin
    SetControlVisible('GPP_TAILLAGEAUTO',false);
  end;
  {$ENDIF}
END;

{$IFDEF GPAO}
procedure TOM_ParPiece.GPCacheChamps;
{ Paramètrage de s pièces GP => cache les champs qui ne sont pas utilisés }
var
	iPage,iControl: Integer;
   C: TControl;
   PC: TPageControl;
begin
	{ Balaye les onglets du PageControl }
   PC := TPageControl(GetControl('PPAGES'));
   if PC <> nil then
   begin
	   for iPage := 0 to PC.PageCount - 1 do
      begin
      	if PC.Pages[iPage].Tag <> 120 then
         begin
            { Balaye les contrôles du TabSheet }
            for iControl := 0 to PC.Pages[iPage].ControlCount - 1 do
            begin
               C := tControl(PC.Pages[iPage].Controls[iCOntrol]);
               if C.Visible then
                 C.Visible := C.Tag = 120;
            end;
         end
         else
            PC.Pages[iPage].TabVisible := False;
      end;
   end;
end;
{$ENDIF}

{$IFDEF BTP}
procedure TOM_ParPiece.ChangeAffDefOuv( Sender : Tobject)  ;
var GestionLib : Integer;
BEGIN
GestionLib := GetField('GPP_TYPEPRESENT');
ParamDetailOuvrage (GestionLib);
// Verrue
DS.edit;
// ------
SetField ('GPP_TYPEPRESENT',GestionLib);
THRichEditOle(GetControl('TMODEAFFDEF')).Lines.text := RecupModeAffLibelles (GestionLib);
END;
{$ENDIF}

//AC 18/08/03 NV GESTION COMPTA DIFF
procedure TOM_ParPiece.TstComptaStock;
var ComptaStock: boolean;
begin
  //Si compta différée dans ParamSoc ou pas d'écritures normales, désactive maj compta stock
  if (GetParamSoc('SO_GCDESACTIVECOMPTA')) or (GetControlText('GPP_TYPEECRCPTA') <> 'NOR') then
    ComptaStock := False
  else
    ComptaStock := True;
  SetControlProperty('GPP_TYPEECRSTOCK', 'Enabled', ComptaStock);
  SetControlProperty('GPP_COMPSTOCKLIGNE', 'Enabled', ComptaStock);
  SetControlProperty('GPP_COMPSTOCKPIED', 'Enabled', ComptaStock);
end;

procedure TOM_ParPiece.ClickPassageCpta(Sender: Tobject);
begin
  TstComptaStock;
end;
//Fin AC

procedure TOM_ParPiece.OnClose  ;
begin
if EstModifie then
   begin
{$IFDEF EAGLCLIENT}
   AvertirCacheServer('PARPIECE') ;
{$ENDIF}
    HShowMessage('0;Modification de paramétrage;'+TraduireMemoire(TexteMessage[11])+';E;O;O;O;', '', '') ;
   end;
TOBDetailPiece.Free;
end;


////////////////////////////////////////////////////////////////////////////////
Procedure TOMParPiece_AffectGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value, St_Text, St : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCPARPIECE') then exit;
Indice := Integer (Parms[1]);
{$IFDEF BTP}
St_Plus := 'AND CO_CODE LIKE "4%"';
{$ELSE}
St_Plus := 'AND CO_CODE NOT LIKE "9%"';
{$ENDIF}
St_Value := string (THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(Indice))).Value);
St_Text := string (THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(Indice))).Text);
For i_ind := 1 to 8 do
    BEGIN
    if i_ind = Indice then continue;
    St := string (THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(i_ind))).Value);
    If St <> '' then St_Plus := St_Plus + ' AND CO_CODE <>"' + St + '"';
    END;
THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(Indice))).Plus := St_Plus;
THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(Indice))).Value := St_Value;
THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(Indice))).Text := St_Text ;
END;

Procedure TOMParPiece_ChangeGroup (parms:array of variant; nb: integer ) ;
var F: TForm;
    St_Plus, St_Value : string;
    Indice, i_ind : integer;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'GCPARPIECE') then exit;
Indice := Integer (Parms[1]);
St_Plus := '';
St_Value := string (THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(Indice))).Value);
if St_Value = '' then
    BEGIN
    For i_ind := Indice + 1 to 8 do
        BEGIN
        THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(i_ind))).Enabled := True;
        THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(i_ind))).Value := '';
        THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(i_ind))).Enabled := False;
        THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(i_ind))).Color := clBtnFace;
        END;
    END else
    BEGIN
    if Indice < 8 then
        BEGIN
        THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(Indice + 1))).Enabled := True;
        THValComboBox (F.FindComponent('GPP_IFL'+InttoStr(Indice + 1))).Color := clWindow;
        END;
    END;
END;

Procedure TOMParPiece_TableLibre (parms:array of variant; nb: integer ) ;
var F: TFFicheListe;
    LaTOM : TOM ;
BEGIN
F := TFFicheListe(Longint (Parms[0]));
if (F.Name <> 'GCPARPIECE') then exit;
LaTOM:=TFFicheListe(F).OM;
if (LaTOM is TOM_ParPiece) then
  begin
  if String(Parms[1])='AFFECTE' then TOM_ParPiece(LaTOM).AffecteValeurTableLibre(Integer(Parms[2]))
    else if String(Parms[1])='CHANGE' then TOM_ParPiece(LaTOM).ChangeTableLibre(Integer(Parms[2]))
      else if String(Parms[1])='CLICKOBL' then TOM_ParPiece(LaTOM).ClickOblTableLibre(Integer(Parms[2]));
  end;
END;

procedure InitTOMParPiece() ;
begin
RegisterAglProc('ParPiece_AffectGroup', True , 1, TOMParPiece_AffectGroup);
RegisterAglProc('ParPiece_ChangeGroup', True , 1, TOMParPiece_ChangeGroup);
RegisterAglProc('ParPiece_TableLibre', True , 2, TOMParPiece_TableLibre);
end;

{$IFDEF GPAO}
procedure TOM_ParPiece.SetConfigState;
begin
  SetControlEnabled('GPP_CFGART', Ecran.Name = 'GCPARPIECE');
  if GetControlEnabled('GPP_CFGART') then
     SetControlEnabled('GPP_CFGARTASSIST', GetCheckBoxState('GPP_CFGART') = cbChecked)
  else
     SetControlEnabled('GPP_CFGARTASSIST', False);
end;
{$ENDIF}


procedure TOM_ParPiece.ParamSaisie(Sender: Tobject);
var ListeSaisie : string;
begin
  ListeSaisie := GetControlText('GPP_LISTESAISIE');
  AGLLanceFiche('BTP','BTCRESAISIE','','','LISTE='+ListeSaisie)
end;


Initialization
registerclasses([TOM_ParPiece,TOM_DomainePiece]);
InitTOMParPiece();
end.
