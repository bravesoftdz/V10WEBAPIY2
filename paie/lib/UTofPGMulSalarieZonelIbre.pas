{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 22/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGANALYSEZONELIBRE ()
Mots clefs ... : TOF;PGANALYSEZONELIBRE
*****************************************************************
PT1     28/09/2007 FC V_80 FQ 14807 Gestion libellé/code
}
Unit UTofPGMulSalarieZonelIbre ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,

{$ENDIF}
     forms,
     uTob,
     P5def,
     EntPaie,
     PGOutils2,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     PgOutilsHistorique,
     HTB97,
     UTOF ; 

Type
  TOF_PGMULSALZONELIBRE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    sMode : String; //PT1
    procedure BMajClick(Sender : TObject);
    procedure ExitEdit(Sender: TObject);
    procedure OnClickSalarieSortie(Sender: TObject);
  end ;

Implementation

procedure TOF_PGMULSALZONELIBRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSALZONELIBRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSALZONELIBRE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSALZONELIBRE.OnLoad ;
begin
  Inherited ;
  //DEB PT1
  if (GetControlText('CKLIBELLE') <> sMode) then
  begin
    if (GetControlText('CKLIBELLE') = '-') then
      PGMajDonneesAfficheZL(False,0,True)
    else
      PGMajDonneesAfficheZL(False,0,False);
    sMode := GetControlText('CKLIBELLE');
  end;
  //FIN PT1
end ;

procedure TOF_PGMULSALZONELIBRE.OnArgument (S : String ) ;
var
  iTable, iChamp: integer;
  St, Prefixe, TheEtab, LesEtab: string;
  Q: TQuery;
  TOB_DesChamps: TOB;
  T1: TOB;
  zz, i: Integer;
  LeType, LaValeur, LePlus, LeChamp: string;
  Bt : TToolBarButton97;
  Defaut: THEdit;
  Check : TCheckBox;
  Num : Integer;
begin
   Inherited;
   Prefixe := 'PTZ';
   iTable := PrefixeToNum(Prefixe);
   ChargeDeChamps(iTable, Prefixe);
   Q := OpenSQL('SELECT * FROM PGPARAMAFFICHEZL',True);
   for iChamp := 1 to high(V_PGI.DeChamps[iTable]) do
   begin
    For i := 1 to 30 do
    begin
      if V_PGI.DEChamps[iTable, iChamp].Nom = 'PTZ_PGVALZL'+IntToStr(i) then
      begin
        LaValeur := Q.FindField('PAZ_CHAMPDISPO'+IntToStr(i)).AsString;
        LeChamp := Copy(LaValeur,4,Length(LaValeur));
        LeType := Copy(LaValeur,1,3);
        If LeType = 'NAT' then
        begin
          V_PGI.DEChamps[iTable, iChamp].Libelle := RechDom('PGELEMENTNAT',LeChamp,False);
          V_PGI.DEChamps[iTable, iChamp].Control := 'LDC';
        end
        else If LeType = 'ZLS' then
        begin
          V_PGI.DEChamps[iTable, iChamp].Libelle := RechDom('PGZONEHISTOSAL',LeChamp,False);
          V_PGI.DEChamps[iTable, iChamp].Control := 'LDC';
        end
        else
        begin
          V_PGI.DEChamps[iTable, iChamp].Control := '';
          V_PGI.DEChamps[iTable, iChamp].Libelle := 'Champ libre non paramétré';
        end;
      end;
    end;
   end;
   Ferme(Q);
   Bt := TToolBarButton97(GetControl('BMAJDONNEES'));
   If Bt <> Nil then Bt.OnClick := BMajClick;
   SetControlvisible('DATEARRET',True);
   SetControlvisible('TDATEARRET',True);
   SetControlEnabled('DATEARRET',False);
   SetControlEnabled('TDATEARRET',False);
   Check:=TCheckBox(GetControl('CKSORTIE'));
   if Check<>nil then Check.OnClick:=OnClickSalarieSortie;
   For Num := 1 to VH_Paie.PGNbreStatOrg do
   begin
     if Num >4 then Break;
     VisibiliteChampSalarie (IntToStr(Num),GetControl ('PSA_TRAVAILN'+IntToStr(Num)),GetControl ('TPSA_TRAVAILN'+IntToStr(Num)));
   end;
  VisibiliteStat (GetControl ('PSA_CODESTAT'),GetControl ('TPSA_CODESTAT')) ; //PT3
  Defaut:=ThEdit(getcontrol('PSA_SALARIE'));
  If Defaut<>nil then Defaut.OnExit:=ExitEdit;
end ;

procedure TOF_PGMULSALZONELIBRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSALZONELIBRE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSALZONELIBRE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGMULSALZONELIBRE.BMajClick(Sender : TObject);
begin
  //DEB PT1
  //  PGMajDonneesAfficheZL;
  if (GetControlText('CKLIBELLE') = '-') then
    PGMajDonneesAfficheZL(False,0,True)
  else
    PGMajDonneesAfficheZL(False,0,False);
  sMode := GetControlText('CKLIBELLE');
  //FIN PT1
end;

procedure TOF_PGMULSALZONELIBRE.OnClickSalarieSortie(Sender: TObject);
begin
SetControlenabled('DATEARRET',(GetControltext('CKSORTIE')='X'));
SetControlenabled('TDATEARRET',(GetControltext('CKSORTIE')='X'));
end;

procedure TOF_PGMULSALZONELIBRE.ExitEdit(Sender: TObject);
var edit : thedit;
begin
edit:=THEdit(Sender);
if edit <> nil then	//AffectDefautCode que si gestion du code salarié en Numérique
    if (VH_Paie.PgTypeNumSal='NUM') and (length(Edit.text)<11) and (isnumeric(edit.text)) then
    edit.text:=AffectDefautCode(edit,10);
end;


Initialization
  registerclasses ( [ TOF_PGMULSALZONELIBRE ] ) ;
end.

