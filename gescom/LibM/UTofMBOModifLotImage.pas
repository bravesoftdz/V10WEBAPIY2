unit UTofMBOModifLotImage;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,Messages,HStatus,GCMzsUtil,UtilArticle, 
{$IFDEF EAGLCLIENT}
      Maineagl,eMul,
{$ELSE}
      Fiche,mul, DBGrids, Fe_Main,HDB,db,dbTables,UTofMulImage,
{$ENDIF}
      M3VM , M3FP,utilGC
       ;

Type

     TOF_MBOModifLotImage = Class (TOF_MULIMAGE)
        CCV : THValComboBox ;
        TypeNomen : THValComboBox;
        procedure OnArgument (S : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure ModificationParLot ;
       private
{$IFDEF BTP}
       procedure ChangeTypeArt(Sender: Tobject);
{$ENDIF}
     END ;

procedure AGLModificationParLot2(parms:array of variant; nb: integer ) ;


implementation

// Modif BTP
procedure TOF_MBOModifLotImage.OnArgument (S : String );
begin
inherited;
{$IFDEF BTP}
CCV := THValComboBox (ecran.FindComponent('GA_TYPEARTICLE'));
CCV.Vide := false;
CCV.ReLoad;
CCV.value := 'MAR';
CCV.OnChange := ChangeTypeArt;
{$ENDIF}
MajChampsLibresArticle(TForm(Ecran));
end;

{$IFDEF BTP}
procedure TOF_MBOModifLotImage.ChangeTypeArt ( Sender : Tobject);
begin
if (CCV.Value = 'OUV') or (CCV.value = 'NOM')   then
   begin
   TypeNomen := THValComboBox (ecran.findComponent('GA_TYPENOMENC'));
   TypeNomen.visible := true;
   TypeNomen.DataType := 'BTTYPENOMENC';
   TypeNomen.Vide := false;
   TypeNomen.ReLoad ;
   if CCV.Value = 'OUV' then TypeNomen.Value := 'OUV'
                        else TypeNomen.Value := 'MAC';
   GetControl ('TGA_TYPENOMENC').Visible := true;
   end else
   begin
   TypeNomen := THValComboBox (ecran.findComponent('GA_TYPENOMENC'));
   TypeNomen.visible := false;
   TypeNomen.DataType := 'GCTYPENOMENC';
   TypeNomen.Vide := true;
   TypeNomen.ReLoad ;
   GetControl ('TGA_TYPENOMENC').Visible := false;
   end;
end;
{$ENDIF}
// -------------

procedure TOF_MBOModifLotImage.OnLoad;
begin
//inherited;
Ecran.Caption:='Saisie en lot d''article';
end;


procedure TOF_MBOModifLotImage.OnUpdate;
var  F : TFMul ;
    i:integer;
    CC:THLabel ;
begin
inherited ;
if not (Ecran is TFMul) then exit;
F:=TFMul(Ecran) ;

for i:=1 to 3 do
    begin
    CC:= THLabel(TFMul(F).FindComponent('TGA_FAMILLENIV'+InttoStr(i)));
    CC.Caption:=RechDom('GCLIBFAMILLE','LF'+InttoStr(i),false);
    end;
{$IFDEF EAGLCLIENT}
// AFAIREEAGL
{$ELSE}
for i:=0 to F.FListe.Columns.Count-1 do
    begin
    if copy(F.FListe.Columns[i].Title.caption,1,14)='Famille niveau' then
        begin
        CC:=THLabel(F.FindComponent('TGA_FAMILLENIV'+copy(F.FListe.Columns[i].Title.caption,16,1))) ;
        F.Fliste.columns[i].visible:=CC.visible ;
        F.Fliste.columns[i].Field.DisplayLabel:=CC.Caption ;
        end;
    end;
{$ENDIF}
end;

/////////////// ModificationParLotDesTiers //////////////
procedure TOF_MBOModifLotImage.ModificationParLot;
Var F : TFMul ;
    ArticleSpecif: String ;
    Parametrages : String ;
    TheModifLot : TO_ModifParLot;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin MessageAlerte('Aucun élément sélectionné'); exit; end;
//if HShowMessage('0;Confirmation;Confirmez vous la modification des clients?;Q;YN;N;N;','','')<>mrYes then exit;
TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe; TheModifLot.Q := F.Q;
TheModifLot.Titre := 'Modification par lot des Articles';
TheModifLot.FCode := 'GA_ARTICLE';
TheModifLot.TableName := 'ARTICLE';
//
// Modif BTP
//
{$IFDEF BTP}
   TheModifLot.Nature := 'BTP';
   if (CCV.value = 'NOM') or (CCV.value = 'OUV') then TheModifLot.StContexte := TypeNomen.Value
                                                 else TheModifLot.StContexte := CCV.Value;
   if CCV.Value = 'POU' Then TheModifLot.FicheAOuvrir := 'BTARTPOURCENT'
                        else if (CCV.Value = 'PRE') then
                                TheModifLot.FicheAOuvrir := 'BTPRESTATION'
                        else TheModifLot.FicheAOuvrir := 'BTARTICLE';
{$ELSE}
// Nom de la fiche à ouvrir en fontion de V_PGI.PGIContexte
ArticleSpecif:=IsArticleSpecif('MAR') ;
if ArticleSpecif='FicheAffaire' then
   begin
   TheModifLot.Nature := 'AFF';
   TheModifLot.FicheAOuvrir := 'AFARTICLE';
   end
else if ArticleSpecif='FicheModeArt' then
   begin
   TheModifLot.Nature := 'MBO';
   TheModifLot.FicheAOuvrir := 'ARTICLE';
   end
else
   begin
   TheModifLot.Nature := 'GC';
   TheModifLot.FicheAOuvrir := 'GCARTICLE';
   end;
{$ENDIF}

ModifieEnSerie(TheModifLot, Parametrages) ;

//F.FListe.AllSelected := False;
//F.bSelectAll.Down := False ;
if F.bSelectAll.Down then //F.bSelectAllClick(nil);
  F.bSelectAll.Click;
end;

/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLModificationParLot2(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTof else exit;
if (TOF_MULIMAGE(TOTOF) is TOF_MBOModifLotImage) then
 TOF_MBOModifLotImage(TOTOF).ModificationParLot
else PgiInfo('Pas trouvé',F.Name);

end;


Initialization
registerclasses([TOF_MBOModifLotImage]);
RegisterAglProc('ModificationParLot2',TRUE,2,AGLModificationParLot2);
end.




