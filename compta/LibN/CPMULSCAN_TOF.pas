unit CPMULSCAN_TOF;

Interface

Uses
     Classes,
     UTof,
     UTob,
     HCtrls,
     Menus,           // PopUpMenu
     Grids,           // TGridDrawState
     Graphics,
{$IFDEF EAGLCLIENT}
     MainEAgl,        // AGLLanceFiche
     UtileAgl, {LanceEtatTob}
{$ELSE}
     Fe_Main,         // AGLLanceFiche
     Db,
     EdtREtat, {LanceEtatTob}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     ULibEcriture ,
     uTofViergeMul;   // Fiche Ancetre Vierge MUL

type

  TOF_CPMULSCAN = Class (TOF_ViergeMul)
   procedure InitPopUp ;
   procedure OnClickGenererArbo ( Sender : TObject) ;
   procedure OnClickHPUP ( Sender : TObject) ;
   procedure OnClickHPDOWN ( Sender : TObject) ;
   procedure OnClickGenererXML ( Sender : TObject) ;
   private
   protected
    procedure RemplitATobFListe ; override ;
    procedure InitControl ; override ;
    function  AjouteATobFListe( vTob : Tob) : Boolean; override ;
    procedure OnPopUpPopF11( Sender : TObject ); override;
   public
    procedure OnLoad                   ; override ;
    procedure OnArgument              (S : String ) ; override ;
    procedure OnClose ; override ;
  end ;

procedure CPLanceFiche_CPMULSCAN ;


implementation

uses
 Controls ,
 SaisComm , // pour le GetO
 ULibScanExpert,
 hmsgbox ,
 StdCtrls , // pour le cbChecked
 paramSoc ,
 SysUtils ;

const
 cFI_TABLE    = 'CPSCANEXPERT' ;

{ TOF_CPMULSCAN }

procedure CPLanceFiche_CPMULSCAN ;
begin
 AGLLanceFiche('CP','CPSCANEXPERT', '', '', '') ;
end ;

function TOF_CPMULSCAN.AjouteATobFListe(vTob: Tob) : Boolean;
begin
 if vTOB.GetString('BOSCANEWS') = '' then
  vTOB.PutValue('BOSCANEWS' , '-') ;
 if vTOB.GetString('BOHP') = '' then
  vTOB.PutValue('BOHP' , '-') ;
 if vTOB.GetString('BOHP') <> '-' then
  vTOB.AddChampSupValeur('NBFICHHP',CScanNbFichierD(vTOB.GetString('DOS_NODOSSIER') ) ) ;
 result := true ;
end;


procedure TOF_CPMULSCAN.InitControl;
begin

 inherited;

 PageControl.ActivePage := PageControl.Pages[0] ;
 FListe.ColTypes[3]     := 'B' ;
 FListe.ColTypes[4]     := 'B' ;
 FListe.ColTypes[5]     := 'B' ;
 
end;

procedure TOF_CPMULSCAN.OnArgument(S: String);
begin

 FFI_Table       := cFI_TABLE;
 FStListeParam   := 'DPSCANEXPERT' ;

 inherited ;

 InitPopUp ;

 Ecran.PopupMenu := PopF11 ;

end;

procedure TOF_CPMULSCAN.OnClose;
begin
  inherited;
end;

procedure TOF_CPMULSCAN.OnLoad;
begin
  inherited;
end ;


procedure TOF_CPMULSCAN.InitPopUp ;
var
 lPopUp : TPopUpMenu ;
 i      : integer ;
begin

 if csDestroying in Ecran.ComponentState then Exit ;

 lPopUp := TPopUpMenu(GetControl('PopUpOutil', True)) ;

 for i := 0 to lPopUp.Items.Count - 1 do
  begin

   if lPopUp.Items[i].Name = 'GENERERARBO' then
    begin
     lPopUp.Items[i].OnClick := OnClickGenererArbo ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'HPUP' then
    begin
     lPopUp.Items[i].OnClick := OnClickHPUP ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'HPDOWN' then
    begin
     lPopUp.Items[i].OnClick := OnClickHPDOWN ;
     Continue;
    end; // if

   if lPopUp.Items[i].Name = 'GENERERXML' then
    begin
     lPopUp.Items[i].OnClick := OnClickGenererXML ;
     Continue;
    end; // if

  end ; // for

end ;


procedure TOF_CPMULSCAN.OnClickHPUP ( Sender : TObject) ;
var
 lTOB : TOB ;
 i    : integer ;
begin

 if PGIask('Confirmez-vous l''activation du MFP3035 ? ' , 'Activation') <> mrYes then exit ;

 if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTob := GetO(FListe,i) ;
      CSetParamScanHP('X',lTOB.GetString('dos_nodossier') ) ;
     end ; // if
  end // if
   else
    begin
     lTob := GetO(FListe) ;
     CSetParamScanHP('X',lTOB.GetString('dos_nodossier') ) ;
    end ;

 RefreshPclPge ;

end ;

procedure TOF_CPMULSCAN.OnClickHPDOWN ( Sender : TObject) ;
var
 lTOB : TOB ;
 i    : integer ;
begin

 if PGIask('Confirmez-vous la désactivation du MFP3035 ? ' , 'Désactivation') <> mrYes then exit ;

 if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTob := GetO(FListe,i) ;
      CSetParamScanHP('-',lTOB.GetString('dos_nodossier') ) ;
     end ; // if
  end // if
   else
    begin
     lTob := GetO(FListe) ;
     CSetParamScanHP('-',lTOB.GetString('dos_nodossier') ) ;
    end ;

 RefreshPclPge ;

end ;

procedure TOF_CPMULSCAN.OnClickGenererXML ( Sender : TObject) ;
var
 lTOB : TOB ;
 i    : integer ;
begin

 if PGIask('Confirmez-vous la génération du fichier de configuration du scanner ' , 'Génération') <> mrYes then exit ;

 if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTob := GetO(FListe,i) ;
      CGenereXML(lTOB.GetString('dos_nodossier') ) ;
     end ; // if
  end // if
   else
    begin
     lTob := GetO(FListe) ;
     CGenereXML(lTOB.GetString('dos_nodossier') ) ;
    end ;

 RefreshPclPge ;

end ;


procedure TOF_CPMULSCAN.OnClickGenererArbo ( Sender : TObject) ;
var
 i    : integer ;
 lTob : TOB ;
 lInResult : integer ;
begin

 if PGIask('Confirmez-vous la création des répertoires pour Cegid Scan Expert ? ' , 'Création des répertoire') <> mrYes then exit ;

 if ABoMultiSelected then
  begin
   for i := 1 to FListe.RowCount - 1 do
    if FListe.IsSelected(i) then
     begin
      lTob := GetO(FListe,i) ;
      lInResult := CCreationArboDoss(lTOB.GetString('dos_nodossier')) ;
      if lInResult = 0 then
       begin
        PGIInfo('Echec à la création du repertoire'+#10#13+GetLastError) ;
        exit ;
       end
        else
         CSetParamScanHP('X',lTOB.GetString('dos_nodossier') ) ;
     end ; // if
  end // if
   else
    begin
     lTob := GetO(FListe) ;
     lInResult := CCreationArboDoss(lTOB.GetString('dos_nodossier')) ;
     if lInResult = 0 then
      begin
       PGIInfo('Echec à la création du repertoire'+#10#13+GetLastError) ;
       exit ;
      end
       else
        CSetParamScanHP('X',lTOB.GetString('dos_nodossier') ) ;
    end ;

  PGIInfo('Création des répertoires réussie') ;

  RefreshPclPge ;

end ;

procedure TOF_CPMULSCAN.RemplitATobFListe ;
var
 lStWhere,lStWhere1,lStWhere2 : string ;
 i : integer ;
begin

 inherited ;

 if GetCheckBoxState('EWS_') = cbChecked then
  lStWhere  :=  ' and dos_ewscree="X"'
   else
    if GetCheckBoxState('EWS_') = cbUnChecked then
     lStWhere  :=  ' and dos_ewscree="-"' ;

 if GetCheckBoxState('SCANEWS_') = cbChecked then
  lStWhere1  :=  ' and yds_libelle="X"'
   else
    if GetCheckBoxState('SCANEWS_') = cbUnChecked then
     lStWhere1  :=  ' and ( (yds_libelle="-") or (yds_libelle is null) ) ' ; //' and yds_libelle="-"' ;

 if GetCheckBoxState('SCANHP_') = cbChecked then
  lStWhere2  :=  ' and yds_libre="X"'
   else
    if GetCheckBoxState('SCANHP_') = cbUnChecked then
     lStWhere2  :=  ' and ( ( yds_libre="-") or ( yds_libre is null) ) ' ;


 AStSqlTobFListe := 'select dos_nodossier,dos_libelle,dos_ewscree,yds_libelle BOSCANEWS,yds_libre BOHP from dossier ' +
                     'left outer join choixdpstd on ( yds_nodossier = dos_nodossier and yds_type="CES" ) ' +
                     ' where dos_cabinet<>"X" and dos_nodossier<>"000STD" '
                     + lStWhere + lStWhere1 + lStWhere2 ;


end;


procedure TOF_CPMULSCAN.OnPopUpPopF11 (Sender: TObject);
var
 lMenuItem : TMenuItem;
begin

 inherited ;

end;



Initialization
  registerclasses ( [ TOF_CPMULSCAN ] ) ;

end.
