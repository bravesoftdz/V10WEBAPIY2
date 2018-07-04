unit UTofAffaire_REg;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,HTB97,
{$IFDEF EAGLCLIENT}
   eMul,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, Mul,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOF, AffaireUtil,utob,grids, Dicobtp, WIndows,
      EntGc,saisutil,vierge,paramsoc{,afprepfact},uEntCommun;
Type
     TOF_AFFAIRE_REG = Class (TOF)
     public
     		procedure OnArgument(stArgument : String ) ; override ;

     private
        GS : THGrid;
        Statut,LesCol : String;
        nbCol, ColAff0,ColAff1,ColAff2,ColAff3,ColAff,Colavnt,ColLib,ColPrin : integer;
        ColDeb,ColFin,ColReg,ColEsc : Integer ;
        RegSurCaf : String;


        procedure BZoomClick ( sender : TObject);
        procedure BValiderClick ( sender : TObject);
        procedure AfficheGrid;
     END ;

implementation

procedure TOF_AFFAIRE_REG.OnArgument(stArgument : String );
Var
	Critere, Champ, valeur  : String;
   ztiers,zreg: String;
   Action : TActionFiche;
   x : Integer;
   Bt : TToolBarButton97;
   st,lecaf,CodeAff: string;
   sttiers,strep: string;
   CleDocTemp : R_CleDoc;
   Nbpiece : Integer;

begin
Inherited;
 // Recup des critères
 if (stArgument = '') then exit;
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'TIERS' then Ztiers := Valeur;
        if Champ = 'REGROUPE' then Zreg := Valeur;
        if Champ = 'STATUTAFFAIRE' then Statut := Valeur;
        if Champ = 'EXER' then SetControlText('AFF_AFFAIRE2', Valeur);
        if Champ = 'LECAF' then LeCAf := Valeur;
        if Champ = 'AFF_REGSURCAF' then REgSurCaf := Valeur;
        if Champ = 'AFF_AFFAIRE' then CodeAff := Valeur;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;

    NbPiece:=0;
    // Recherche du CAF
    if (RegSurCaf='X') and (LeCaf = '') then
				NbPiece := SelectPieceAffaireBis(CodeAff, 'AFF', CleDocTemp ,Sttiers,StRep);
    		if (NbPiece =1) then  LeCaf := Sttiers;
    Begin

    End;

if (RegsurCaf='X') then
  Begin
    SetControlText('AFF_TIERS',LeCaf);
    SetControlText('TREG_TIERS','Client Facturé');
  End
else
  Begin
  SetControlText('AFF_TIERS',ztiers);
  SetControlText('TREG_TIERS','Client ');
  End;

SetControlText('AFF_REGROUPEFACT',zreg);
if (GetParamSoc('SO_AfFormatExer') = 'AUC') then begin
   SetControlVIsible ('TREG_AFFAIRE2',false);
   SetControlVIsible ('AFF_AFFAIRE2',false);
   SetControlVIsible ('BEXER',false);
   end;
// Zoom Click
Bt := TToolBarButton97 (GetControl('BZOOM'));
if Bt <> Nil then Bt.OnClick := BZoomClick;
Bt := TToolBarButton97 (GetControl('BExer'));
if Bt <> Nil then Bt.OnClick := BValiderClick;

Colavnt:=-1; COlAff3:=-1;Colaff2:=-1;
// Gestion du Grid
GS := THGRID(GetControl('GRID_LISTEAFFREG'));
nbCol := 8; LesCol := 'FIXED;AFF_PRINCIPALE;AFF_AFFAIRE1;';
if (VH_GC.CleAffaire.NbPartie>1) then
   BEGIN Inc(NbCol); LesCol:=LesCol + 'AFF_AFFAIRE2;' END;
if (VH_GC.CleAffaire.NbPartie>2)  then      // obligation d'afficher le compteur même si pas gérer, sinon, le zoom ne marche pas
   BEGIN Inc(NbCol); LesCol:=LesCol + 'AFF_AFFAIRE3;' END;
If (GetParamSoc('SO_AffGestionAvenant')=true) then begin Inc(NbCol);LesCol := LesCol + 'AFF_AVENANT;'; end;
LesCol := LesCol + 'AFF_LIBELLE;AFF_DATEDEBGENER;AFF_DATEFINGENER;MR_LIBELLE;GP_ESCOMPTE';
GS.ColCount:=NbCol;

St:=LesCol;
for x:=0 to GS.ColCount-1 do
   BEGIN
   if x>2 then  GS.ColWidths[x]:=100;
   Champ:=ReadTokenSt(St) ;
   if Champ='AFF_AFFAIRE' then ColAff := x
   else if Champ='AFF_AFFAIRE0' then ColAff0 := x
   else if Champ='AFF_AFFAIRE1' then ColAff1 := x
   else if Champ='AFF_AFFAIRE2' then ColAff2 := x
   else if Champ='AFF_AFFAIRE3' then begin
        ColAff3 := x ;
        if GetParamsoc('SO_AFFCO3VISIBLE')=False then GS.ColWidths[ColAFF3]:=0;
        end
   else if Champ='AFF_PRINCIPALE'  then ColPrin:= x
   else if Champ='AFF_AVENANT'   then ColAvnt:= x
   else if Champ='AFF_LIBELLE'   then ColLib := x
   else if Champ='AFF_AVENANT'   then ColAvnt:= x
   else if Champ='AFF_DATEDEBGENER'   then ColDeb:= x
   else if Champ='AFF_DATEFINGENER'   then ColFin:= x
   else if Champ='MR_LIBELLE'   then ColReg:= x
   else if Champ='GP_ESCOMPTE'   then ColEsc:= x
   ;
   END ;
// libellé des colonnes
GS.ColWidths[ColLib]:=200;
GS.ColWidths[ColDeb]:=130;
GS.ColWidths[ColFin]:=130;
GS.ColWidths[ColReg]:=130;
GS.ColWidths[ColEsc]:=100;
GS.Cells[ColPrin,0]:= 'Princ.';
GS.Cells[ColAff1,0]:= VH_GC.CleAffaire.Co1Lib;
if ColAff2 <> -1 then GS.Cells[ColAff2,0]:= VH_GC.CleAffaire.Co2Lib;
if ColAff3 <> -1 then GS.Cells[ColAff3,0]:= VH_GC.CleAffaire.Co3Lib;
GS.ColWidths[0]:=15;  // Fixed col
GS.Cells[ColLib,0]:= 'Libellé';
if ColAvnt<>-1 then GS.Cells[ColAvnt,0]:= 'Avenant';
GS.Cells[ColDeb,0]:= 'Debut';
GS.Cells[ColFin,0]:= 'Fin';
GS.Cells[ColReg,0]:= 'Reglt';
GS.Cells[ColEsc,0]:= 'Esc.';

//GS.ColWidths[ColAff0]:=0;  // affaire 0
Action := taConsult;
AffecteGrid(GS,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GS) ;
//TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
GS.ColWidths[0]:=15;  // Fixed col

// Chargement de la TOB pour affichage du GRID des affaires regroupées
AfficheGrid;

End;


procedure TOF_AFFAIRE_REG.AfficheGrid;
var
   TobReg : TOB;
   TobDet : TOB;
   QQ : Tquery;
   Req, Nat: String;
   wi : Integer;
begin
if Statut = 'PRO' then Nat := GetParamSoc('SO_AFNATPROPOSITION')
                  else Nat := GetParamSoc('SO_AFNATAFFAIRE');
TobReg := Tob.Create('Liste Affaires Regroupées',nil,-1);
QQ := nil;
Try
Req := 'SELECT AFF_AFFAIRE1,AFF_AFFAIRE2,AFF_AFFAIRE3,AFF_AVENANT,AFF_LIBELLE,AFF_PRINCIPALE'+
',AFF_DATEDEBGENER,AFF_DATEFINGENER,GP_MODEREGLE,MR_LIBELLE,GP_ESCOMPTE'+
' FROM AFFAIRE '+
' left outer join PIECE on GP_AFFAIRE=AFF_AFFAIRE '+
' left outer join MODEREGL on GP_MODEREGLE=MR_MODEREGLE ' +
' Where GP_NATUREPIECEG="'+Nat+'"';

If (RegsurCaf ='X') then
	Req := Req + ' and GP_TIERSFACTURE="'+GetControlText('AFF_TIERS')+ '"' + ' and AFF_REGSURCAF="X"'
else
	Req := Req + ' and AFF_TIERS="'+GetControlText('AFF_TIERS')+ '"';

if (GetParamSoc('SO_AFFormatExer')<>'AUC') and (GetControlText('AFF_AFFAIRE2') <>'')
   then req:=req +'AND AFF_AFFAIRE2="'+GetControlText('AFF_AFFAIRE2')+'"';

req:=req +' AND AFF_REGROUPEFACT ="'+GetControlText('AFF_REGROUPEFACT')+'"order by  AFF_PRINCIPALE DESC';

QQ := OpenSQL(Req,True) ;
If Not QQ.EOF then TobReg.LoadDetailDB('REGROUPE_AFFAIRE','','',QQ,True) Else Exit;
Finally
 Ferme(QQ);
End;
if (TobReg.Detail.Count > 0) Then
Begin
  for wi:=0 To TobReg.Detail.Count-1 do
  Begin
    TobDet := TobReg.Detail[wi];
    If (TobDet.Getvalue('AFF_PRINCIPALE')= 'X') then   // Aff. Princ.
          TobDet.PutValue('AFF_PRINCIPALE','  X  ')
    else
          TobDet.PutValue('AFF_PRINCIPALE',' ');
   End;   // for
End;

TobReg.PutGridDetail(GS,False,False,LesCol,true);
TobReg.free;
end;

procedure TOF_AFFAIRE_REG.BZoomClick ( sender : TObject);
Var CodeAffaire,Aff0,Aff1, Aff2, Aff3, Aff4 : string;
    row : integer;
begin
Aff0 := ''; Aff1 := ''; Aff2 := ''; Aff3 := ''; Aff4 := '';
row := GS.row;
Aff0 := StatutCompletToReduit(Statut);
if ColAff1 <> 0 then Aff1 := GS.Cells[ColAff1, Row];
if ColAff2 <> -1 then Aff2 := GS.Cells[ColAff2, Row];
if ColAff3 <> -1 then Aff3 := GS.Cells[ColAff3, Row];
if ColAvnt <> -1 then Aff4 := GS.Cells[ColAvnt, Row]; //GISE
CodeAffaire := CodeAffaireRegroupe(Aff0,Aff1, Aff2, Aff3, Aff4, taModif, false,false,false);
if not ExisteAffaire (CodeAffaire,'') then BEGIN PGIBoxAF ('Affaire non valide',''); Exit; END
else V_PGI.DispatchTT( 5, taModif, CodeAffaire, '', 'MONOFICHE') ;
end;

procedure TOF_AFFAIRE_REG.BValiderClick ( sender : TObject);
begin
if GetParamSoc('SO_AFFormatExer')<>'AUC' then begin
   AfficheGrid;
   SETfocusControl('AFf_AFFAIRE2');
   exit;
   end;
end;

Initialization
registerclasses([TOF_AFFAIRE_REG]);
end.
