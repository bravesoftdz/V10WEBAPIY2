unit UtofPtfPieceCegid;

interface
uses  M3FP, StdCtrls,Controls,Classes, Graphics, forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOB,uTofAfBaseCodeAffaire,
      
{$IFDEF EAGLCLIENT}
      MaineAGL,eQRS1,
{$ELSE}
      Fe_Main,db,dbTables,MajTable,QRS1,
{$ENDIF}
      UTOF, AglInit, Agent, Ent1 ;

Type
     TOF_PtfPieceCegid = Class (TOF_AFBASECODEAFFAIRE)
       private
       stNatureEtat : string ;

       public
       procedure OnArgument (stArgument : String )  ; override ;
       procedure OnNew ; override;
       procedure OnUpdate ; override ;
       procedure OnLoad ; override ;
     END ;

Const
	// libellés des messages
	TexteMessage: array[1..1] of string 	= (
          {1}         'Vous devez sélectionner une nature de pièce.'
                 );
implementation

/////////////////////////////////////////////////////////////////////////////
procedure TOF_PtfPieceCegid.OnArgument (stArgument : String ) ;
Var St_Plus, stArg, stVenteAchat : string ;
    CC : THValComboBox ;
    Iind : integer;
begin
    Inherited;
stArg := stArgument;
stVenteAchat := ReadTokenSt(stArg); stNatureEtat := ReadTokenSt(stArg);
if stVenteAchat='VENTE' then St_Plus:=' AND GPP_VENTEACHAT="VEN"'
else if stVenteAchat='ACHAT' then St_Plus:=' AND GPP_VENTEACHAT="ACH"'
else St_Plus:=' AND GPP_VENTEACHAT="AUT"' ;

//*** Blocages des natures de pièces autorisées en affaires
if (ctxAffaire in V_PGI.PGIContexte) and (stVenteAchat='ACHAT') then
   begin
   if ctxScot in V_PGI.PGIContexte then
      begin
      SetControlText ('NATUREPIECEG','FF');
      Setcontrolenabled ('NATUREPIECEG',False);
      end
   else
      begin
      st_plus:=st_Plus + ' AND ((GPP_NATUREPIECEG="CF") or (GPP_NATUREPIECEG="BLF") or (GPP_NATUREPIECEG="FF"))';
      end;
   end;
// **** Fin affaire ***

if not(ctxScot in V_PGI.PGIContexte) then THValComboBox (GetControl('NATUREPIECEG')).Plus := St_Plus ;
THValComboBox (GetControl('GP_DEPOT')).ItemIndex := 0;
THValComboBox (GetControl('GP_ETABLISSEMENT')).ItemIndex := 0;

if stVenteAchat='VENTE' then
   begin
   Iind:=THValComboBox(GetControl ('NATUREPIECEG')).Values.IndexOf('FAC');
   if Iind<0 then Iind:=0;
   THValComboBox(GetControl ('NATUREPIECEG')).Items.Insert(Iind,'Facture + Avoir clients');
   THValComboBox(GetControl ('NATUREPIECEG')).Values.Insert(Iind,'ZZ1');
   SetControlText('NATUREPIECEG','ZZ1');
   SetControlProperty('GP_TIERS','DataType','GCTIERSCLI') ;
   end else if stVenteAchat='ACHAT' then
   begin
   if not(ctxScot in V_PGI.PGIContexte) then
      SetControlText('NATUREPIECEG','FF');
   SetControlProperty('GP_TIERS','DataType','GCTIERSFOURN') ;
   SetControlProperty('GP_TIERS_','DataType','GCTIERSFOURN') ;
   end;

CC:=THValComboBox(GetControl('GP_DOMAINE')) ; if CC<>Nil then PositionneDomaineUser(CC) ;
CC:=THValComboBox(GetControl('GP_ETABLISSEMENT')) ; if CC<>Nil then PositionneEtabUser(CC) ;
end;

procedure TOF_PtfPieceCegid.OnNew;
begin
Inherited;
if stNatureEtat='' then exit;
SetControlText('FEtat',stNatureEtat);
SetControlEnabled('FEtat',False);
TFQRS1(Ecran).FCodeEtat:=GetControlText('FEtat');
end;

procedure TOF_PtfPieceCegid.OnLoad;
var stWhere : string;
begin
stWhere := '';
if stNatureEtat<>'' then SetControlText('FEtat',stNatureEtat);    // PCS 17/03/2003
    Inherited;
if GetControlText('NATUREPIECEG') = '' then
    BEGIN
    SetControlText('XX_WHERE', 'GP_NATUREPIECEG=""');
    LastError:=1;
    LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    END ELSE
    BEGIN
    if GetControlText('NATUREPIECEG')='ZZ1' then
       stWhere := stWhere + ' AND (GP_NATUREPIECEG="FAC" OR GP_NATUREPIECEG="AVC" OR GP_NATUREPIECEG="AVS")'
       else
       stWhere := stWhere + ' AND GP_NATUREPIECEG="' + GetControlText ('NATUREPIECEG') + '"';
    SetControlText('XX_WHERE',stWhere);
    END;
    
if Ecran.Name = 'GCPTFPIECECEGID01' then
   begin
   SetControlText('GL_ARTICLE',GetControlText ('CODEARTICLE'));
   if GetControlText ('CODEARTICLE_') <> '' then
       SetControlText('GL_ARTICLE_',format('%-18.18s%16.16s',[GetControlText ('CODEARTICLE_'),StringOfChar('Z',16)]))
   else SetControlText('GL_ARTICLE_','');
   end;
end;

procedure TOF_PtfPieceCegid.OnUpdate;
begin
    Inherited;
if GetControlText('NATUREPIECEG') = '' then
    BEGIN
    SetControlText('XX_WHERE', '');
    END;
end;


/////////////////////////////////////////////////////////////////////////////

Initialization
registerclasses([TOF_PtfPieceCegid]) ;
end.
