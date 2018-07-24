Unit UTomModeregl;

Interface

Uses  StdCtrls,Classes,
{$IFDEF EAGLCLIENT}
      utob,eFichList,
{$ELSE}
      db,FichList,
{$ENDIF}
      sysutils,
      HCtrls,UTOM;  

Type
        TOM_MODEREGL = class(TOM)
        Private
          Guide : Boolean ;
          CanCreate : boolean;

          procedure SetButtonEna;

        Protected
          procedure OnUpdateRecord ; override ;
          procedure OnChangeField (F : TField)  ; override ;
          procedure OnNewRecord  ; override ;
          procedure OnDeleteRecord ; override ;

          procedure OnArgument (stArgument : String ) ; override ;
          function  ModePaieOk : Boolean ;
          function  TotalTauxOK : Boolean ;
          function  PresenceModeRegle : Boolean ;
End;

Const
	// libellés des messages
	TexteMessage: array[1..8] of string 	= (
          {1}         'Vous devez renseigner un mode de paiement.'
          {2}        ,'Vous devez renseigner un taux de répartition.'
          {3}        ,'Votre saisie est incorrecte : la somme des répartitions ne fait pas 100%.'
          {4}        ,'Vous devez renseigner la date de départ des échéances.'
          {5}        ,'Votre saisie est incorrecte : vous devez renseigner un montant positif.'
          {6}        ,'Vous devez renseigner un code.'
          {7}        ,'Vous devez renseigner un libellé.'
          {8}        ,'Vous ne pouvez pas supprimer ce mode de règlement : il est référencé par un tiers, un général ou dans un guide.'
                );



implementation

uses
  CommonTools
  , HSysMenu
  , HTB97
  ;

procedure TOM_MODEREGL.OnArgument(stArgument : String ) ;
var
  {$IFNDEF EAGLCLIENT}
  x : integer ;
  {$ENDIF !EAGLCLIENT}
  FF : TFFicheListe ;
  {$IFDEF GCGC}
  CB : TCheckBox ;
  {$ENDIF}
begin
  Inherited;
  FF:=TFFicheListe(ecran) ;
  CanCreate := Tools.CanInsertedInTable('MODEREGL'{$IFDEF APPSRV}, '', '' {$ENDIF APPSRV});
  if not CanCreate then
  begin
    FicheReadOnly(FF);
//    FF.TypeAction := taConsult;
  end;
  SetButtonEna;
  {$IFDEF EAGLCLIENT}
  //AFAIREEAGL
  {$ELSE}
  FF.ta.filter:='MR_MODEGUIDE=''-''' ;
  x:=pos('GUIDE',stArgument);
  Guide:=(x>0) ;
  FF.ta.filtered:=(not Guide) ;
  if FF.FLeQuel<>'' then FF.BValider.Enabled:=Guide;
  {$ENDIF}
  {$IFDEF GCGC}
  CB:=TCHeckBox(FF.FindComponent('MR_EINTEGREAUTO')) ;
  if CB<>Nil then CB.Visible:=True ;
  {$ENDIF}
  SetActiveTabSheet('TABCAR');
end ;

procedure TOM_MODEREGL.OnChangeField (F : TField);
var i, iNbEcheances : integer ;
begin
Inherited;
    if ((F.FieldName='MR_NOMBREECHEANCE') and (F.value<1)) then SetField('MR_NOMBREECHEANCE',1) ;
    if (F.FieldName='MR_NOMBREECHEANCE') then
    begin
        iNbEcheances := F.Value;
        For i:=1 to 12 do
        begin
        SetControlVisible('MR_MP'+InttoStr(i),i<=iNbEcheances);
        if not(getcontrol('MR_MP'+InttoStr(i)).visible) then
            if (getfield('MR_MP'+InttoStr(i))<>'') then SetControlText('MR_MP'+InttoStr(i),'') ;
        SetControlVisible('MR_TAUX'+InttoStr(i),i<=iNbEcheances) ;
        if not (getcontrol('MR_TAUX'+InttoStr(i)).visible) then
            if (getfield('MR_TAUX'+IntToStr(i))<>'0') then SetControlText('MR_TAUX'+InttoStr(i),'0');
        // SetControlVisible('MR_ESC'+InttoStr(i),i<=F.value) ; if not TControl(C).Visible then if TDBCheckBox(C).Checked then TDBCheckBox(C).Checked:=False ;
        SetControlVisible('TMR_MP'+InttoStr(i),i<=iNbEcheances) ;

        end ;
    end;
end;


procedure TOM_MODEREGL.OnUpdateRecord ;
var    RP : THDBValComboBox ;
       C : double ;
begin
Inherited;

   SetActiveTabSheet('TabCar') ;
   if (GetField('MR_MODEREGLE')='') then begin SetFocusControl('MR_MODEREGLE'); LastError:=6 ;   LastErrorMsg:=TexteMessage[LastError] ; exit ; end ;

   if (GetField('MR_LIBELLE')='') then begin SetFocusControl('MR_LIBELLE'); LastError:=7 ;   LastErrorMsg:=TexteMessage[LastError] ; exit ; end ;

   if (GetField('MR_APARTIRDE')='') then begin SetFocusControl('MR_APARTIRDE'); LastError:=4 ; LastErrorMsg:=TexteMessage[LastError] ;exit; ; end ;

   if not ModePaieOk then begin LastError:=1 ; LastErrorMsg:=TexteMessage[LastError] ; exit ; end;

   if not TotalTauxOK then begin LastError:=2 ; LastErrorMsg:=TexteMessage[LastError] ; exit ; end;

   RP:=THDBValComboBox(GetControl('MR_REMPLACEMIN')) ;
   if ((RP.Value='') and (RP.Values.Count>0)) then RP.Value:=RP.Values[0] ;

   try
       C:=GetField('MR_MONTANTMIN') ;
   except
       C:=0 ;
   end;
   if (C<0) then begin SetActiveTabSheet('TabCar') ; SetFocusControl('MR_MONTANTMIN') ; LastError:=5 ; LastErrorMsg:=TexteMessage[LastError] ;exit; end ;

end ;


Procedure TOM_MODEREGL.OnNewRecord ;
var i : byte ;
begin
    Inherited ;
    SetField('MR_NOMBREECHEANCE','1') ;
    for i:=1 to 12 do
    begin
        SetField('MR_TAUX'+InttoStr(i),'0') ;
        SetField('MR_ESC'+InttoStr(i),'-') ;
    end ;
  if Guide then SetField('MR_MODEGUIDE','X') else SetField('MR_MODEGUIDE','-');
{$IFDEF EAGLCLIENT}
   if Ecran is TFFicheListe then
      TFFicheListe(Ecran).ta.CurrentFille.PutEcran(TFFicheListe(Ecran)) ;
{$ENDIF}
  SetActiveTabSheet('TABCAR') ;
  SetFocusControl('MR_MODEREGLE');
end ;


procedure TOM_MODEREGL.OnDeleteRecord ;
begin
    Inherited;
    if PresenceModeRegle then Exit ;
end;

function  TOM_MODEREGL.ModePaieOk : Boolean ;
var i : integer ;
begin
{Contrôle si le(s) mode(s)de paiement sont bien renseignés}
     result:=true ;
     for i:=1 to GetField('MR_NOMBREECHEANCE') do
        if (GetField('MR_MP'+InttoStr(i))='') then
        begin
            SetActiveTabSheet('TABECHE') ;
            SetFocusControl('MR_MP'+InttoStr(i));
            result:=false;
            exit;
        end;
end;


function  TOM_MODEREGL.TotalTauxOK : Boolean ;
Var i, j : Byte ;
    Total, C : double ;
begin
    result:=true ;
{ Calcul si le total des taux fait 100% et détruit les anciennes valeurs }
    Total:=0 ;
    for i:=1 to GetField('MR_NombreEcheance') do
    begin
        try
          C:=GetField('MR_TAUX'+InttoStr(i)) ;
        except
          C:=0 ;
        end;
        { 29/03/2006 YMO (FQ 17274-Tâche 3312) Possibilité de saisir 0% 
        if (C <= 0) then
        begin
            SetActiveTabSheet('TABECHE') ;
            SetFocusControl('MR_TAUX'+InttoStr(i)) ;
            result:=false ;
            exit;
        end;
        }
        Total:=Total+C ;
    end ;
    if Abs(Total - 100) < 0.000005 then
    begin
        for j:=GetField('MR_NombreEcheance')+1 to 12 do
            SetField('MR_TAUX'+InttoStr(j),'0') ;
        end
   else
   begin
       SetActiveTabSheet('TABECHE') ;
       LastError:=3 ; LastErrorMsg:=TexteMessage[LastError] ;
       exit ;
   end ;
end;


function TOM_MODEREGL.PresenceModeRegle : Boolean ;
var Trouv : boolean ;
begin
    Result:=False ;
    Trouv:=Presence('TIERS','T_MODEREGLE',GetField('MR_MODEREGLE')) ;
    if not Trouv then Trouv:=Presence('GENERAUX','G_MODEREGLE',GetField('MR_MODEREGLE')) ;
    if not Trouv then Trouv:=Presence('ECRGUI','EG_MODEREGLE',GetField('MR_MODEREGLE')) ;
    if Trouv then begin LastError:=8 ; LastErrorMsg:=TexteMessage[LastError] ; exit ; Result:=True ; end ;
end ;


procedure TOM_MODEREGL.SetButtonEna;
begin
  TtoolbarButton97(GetControl('bInsert')).Visible := CanCreate;
  TtoolbarButton97(GetControl('bDelete')).Visible := CanCreate;
end;

Initialization
registerclasses([TOM_MODEREGL]);

end.
