unit FactCodeBarre;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
{$ENDIF}
{$IFDEF FOS5}
     FOUtil,
{$ENDIF}
     SysUtils, Dialogs, Utiltarif,TarifUtil, UtilPGI, UtilGC, AGLInit, EntGC, SaisUtil, Windows,
     Forms, Classes;

Function FusionnerLignesCodeBarre(TOBPiece : TOB): Boolean;
function QuantiteAutorise(Qte: double): boolean;

implementation

uses FactComm;

function QuantiteAutorise(Qte: double): boolean;
{$IFDEF FOS5}
var QteMax: double;
{$ENDIF}
begin
  Result := True;
  {$IFDEF FOS5}
  QteMax := Arrondi(FOGetParamCaisse('QTEMAX'), V_PGI.OkDecQ);
  Qte := Arrondi(Abs(Qte), V_PGI.OkDecQ);
  if QteMax > 0 then Result := (Qte <= QteMax);
  {$ENDIF}
end;
                         
//Extraction des lignes de codes barres n'ayant pas de lignes génériques dans
//une TOB à part
function ExtraireLesLignesCodeBarre(TOBPiece : TOB): TOB;
var TOBLignesCB : TOB;
    i : integer;
begin
  TOBLignesCB := TOB.Create('lignes code barre',nil,-1);
  For i:=TOBPiece.Detail.Count-1 Downto 0 do
    begin
    if TOBPiece.Detail[i].FieldExists('REGROUPE_CB') then
      begin
      if TOBPiece.Detail[i].GetValue('REGROUPE_CB')='-' then
        begin
        TOBPiece.Detail[i].PutValue('REGROUPE_CB','X');
        TOBPiece.Detail[i].ChangeParent(TOBLignesCB,-1);
        end;
      end;
    end;
  Result:=TOBLignesCB;
end;

procedure RegroupeLignesDistinctes(TOBCB : TOB);
var TOBFusionner,TOBFirst,TOBNext : TOB;
    CodeBarre : String;
    PUHT,ESCOMPTE,REMISE,QTE : Double;
    MTRESTE : Double;
    i : integer;
begin
  TOBFusionner := TOB.Create('Pour la fusion',Nil,-1);
  While TOBCB.Detail.Count>0 do
    begin
    CodeBarre := TOBCB.detail[0].GetValue('GL_REFARTBARRE');
    PUHT := TOBCB.detail[0].GetValue('GL_PUHT');
    ESCOMPTE := TOBCB.detail[0].GetValue('GL_ESCOMPTE');
    REMISE := TOBCB.detail[0].GetValue('GL_REMISELIGNE');
    TOBFirst := TOBCB.FindFirst(['GL_REFARTBARRE','GL_PUHT','GL_ESCOMPTE','GL_REMISELIGNE'],
                                [CodeBarre,PUHT,ESCOMPTE,REMISE],False);
    if TOBFirst<>nil then TOBFirst.ChangeParent(TOBFusionner,-1);
    TOBNext := TOBCB.FindNext(['GL_REFARTBARRE','GL_PUHT','GL_ESCOMPTE','GL_REMISELIGNE'],
                              [CodeBarre,PUHT,ESCOMPTE,REMISE],False);
    while TOBNext<>nil do
      begin
      QTE := Valeur(TOBFirst.GetValue('GL_QTEFACT')) + Valeur(TOBNext.GetValue('GL_QTEFACT'));
      TOBFirst.PutValue('GL_QTEFACT',   QTE);
      TOBFirst.PutValue('GL_QTESTOCK',  QTE);
      TOBFirst.PutValue('GL_QTERESTE',  QTE);
      // --- GUINIER ---
      MTRESTE := Valeur(TOBFirst.GetValue('GL_MONTANTHTDEV')) + Valeur(TOBNext.GetValue('GL_MONTANTHTDEV'));
      TOBFirst.PutValue('GL_MTRESTE',  MTRESTE);
      //
      TOBFirst.Putvalue('GL_RECALCULER','X');
      TOBNext.Free;
      TOBNext:=TOBCB.FindNext(['GL_REFARTBARRE','GL_PUHT','GL_ESCOMPTE','GL_REMISELIGNE'],
                              [CodeBarre,PUHT,ESCOMPTE,REMISE],False);
      end;
    end;
  For i:=TOBFusionner.Detail.Count-1 Downto 0 do TOBFusionner.Detail[i].ChangeParent(TOBCB,-1);
  TOBFusionner.Free;
end;

procedure InsereLigneGeneriqueDansPiece(TOBPiece,TOBLigneCB : TOB);
var TOBL : TOB;
    i : integer;
    CodeGen : string ;
begin
  TOBL := TOB.Create('Une ligne',TOBPiece,-1);
  TOBL.Dupliquer(TOBLigneCB,True,True);

  With TOBL do
    begin
    PutValue('GL_TYPEDIM','GEN'); PutValue('GL_TYPELIGNE','COM');
    PutValue('GL_VALIDECOM','AFF'); PutValue('GL_TYPEREF','');
    PutValue('GL_TENUESTOCK','-');

    CodeGen := GetValue('GL_CODEARTICLE');
    PutValue('GL_CODESDIM',CodeGen); PutValue('GL_REFARTSAISIE',CodeGen) ;
    PutValue('GL_CODEARTICLE',''); PutValue('GL_REFARTBARRE','');
    PutValue('GL_ARTICLE',''); PutValue('GL_TYPEARTICLE','');

    For i:=1 to $A do PutValue('GL_LIBREART'+Format('%x',[i]),'');

    PutValue('GL_REMISABLELIGNE','-'); PutValue('GL_REMISABLEPIED','-');
    PutValue('GL_TOTALHT',0); PutValue('GL_TOTALHTDEV',0); 
    PutValue('GL_TOTALTTC',0); PutValue('GL_TOTALTTCDEV',0); 
    PutValue('GL_MONTANTHT',0); PutValue('GL_MONTANTHTDEV',0); 
    PutValue('GL_MONTANTTTC',0); PutValue('GL_MONTANTTTCDEV',0); 
    For i:=1 to 5 do PutValue('GL_TOTALTAXE'+IntToStr(i),0);
    For i:=1 to 5 do PutValue('GL_TOTALTAXEDEV'+IntToStr(i),0);
    end;
end;

//Parcours toutes les dimensions de l'article générique pour les compter
function IndiceGENPlusDimensions(TOBPiece : TOB; IndiceGEN : Integer) : Integer;
var TOBL : TOB;
    Ligne,NbLignes : integer;
    CodeArticle : string;
    Continue : Boolean;
begin
Result:=-1; NbLignes:=0; Ligne:=IndiceGEN; Continue:=True;
TOBL := TOBPiece.Detail[Ligne]; if TOBL=nil then exit;
if TOBL.GetValue('GL_TYPEDIM')<>'GEN' then exit;
CodeArticle:=TOBL.GetValue('GL_CODESDIM');
Ligne:=Ligne+1;
while Continue and (TOBL<>nil) and (Ligne<TOBPiece.Detail.Count) do
  begin
  TOBL:=TOBPiece.Detail[Ligne];
  if TOBL<>nil then
    begin
    if (TOBL.GetValue('GL_CODEARTICLE')=CodeArticle) and (TOBL.GetValue('GL_TYPEDIM')='DIM') then
      begin Ligne:=Ligne+1; NbLignes:=NbLignes+1; end
    else Continue:=False;
    end;
  end;
Result:=IndiceGEN+NbLignes;
end;

//Parcours toutes les dimensions de l'article générique et recherche une dimension correspondant au code barre
function IsDimensionExiste(TOBPiece : TOB; IndiceGEN : Integer; CodeBarre : String) : boolean;
var TOBL : TOB;
    Ligne : integer;
    CodeArticle : string;
    Continue,Trouve : Boolean;
begin
Result:=False; Ligne:=IndiceGEN; Continue:=True; Trouve:=False;
TOBL := TOBPiece.Detail[Ligne]; if TOBL=nil then exit;
if TOBL.GetValue('GL_TYPEDIM')<>'GEN' then exit;
CodeArticle:=TOBL.GetValue('GL_CODESDIM');
Ligne:=Ligne+1;
while Continue and (TOBL<>nil) and (Ligne<TOBPiece.Detail.Count) do
  begin
  TOBL:=TOBPiece.Detail[Ligne];
  if TOBL<>nil then
    begin
    if (TOBL.GetValue('GL_CODEARTICLE')=CodeArticle) and (TOBL.GetValue('GL_TYPEDIM')='DIM') then
      begin
      if TOBL.GetValue('GL_REFARTBARRE')=CodeBarre then begin Trouve:=True; Continue:=False; end;
      Ligne:=Ligne+1
      end
    else Continue:=False;
    end;
  end;
Result:=Trouve;
end;

Procedure InsereLigneDansPiece(TOBPiece,TOBLigneCB : TOB; Indice : Integer);
var TOBL : TOB;
begin
  TOBL := TOB.Create('Une ligne',TOBPiece,Indice);
  TOBL.Dupliquer(TOBLigneCB,True,True);
  if TOBL.GetValue('UNI_OU_DIM')='DIM' then TOBL.PutValue('GL_TYPEDIM','DIM');
end;

//Insere les lignes de code barre dans la Piece en ajoutant si besoin une ligne
//générique
procedure InsererTOBCBDansTOBPiece(TOBPiece,TOBCB : TOB);
var TOBLigneCB,TOBUniPiece,TOBDimPiece,TOBGENPiece : TOB;
    DimExiste : boolean;
    Indice : Integer;
    QTE : Double;
    MTRESTE : Double;
begin
   While TOBCB.Detail.Count>0 do //Vide la TOB qui a permis la fusion des lignes de Code Barre
     begin
     TOBLigneCB := TOBCB.Detail[0];

     //Insere la ligne si c'est un article unique
     if TOBLigneCB.GetValue('UNI_OU_DIM')='UNI' then
       begin
       //Recherche la ligne DIM identique dans TOBPiece
       TOBUniPiece := TOBPiece.FindFirst(['GL_REFARTBARRE','GL_TYPEDIM','GL_PUHT','GL_ESCOMPTE','GL_REMISELIGNE'],
                      [TOBLigneCB.GetValue('GL_REFARTBARRE'),'NOR',TOBLigneCB.GetValue('GL_PUHT'),
                      TOBLigneCB.GetValue('GL_ESCOMPTE'),TOBLigneCB.GetValue('GL_REMISELIGNE')],False);
       if TOBUniPiece <> Nil then //Dimension trouvée
         begin
         QTE := Valeur(TOBUniPiece.GetValue('GL_QTEFACT'))+Valeur(TOBLigneCB.GetValue('GL_QTEFACT'));
         TOBUniPiece.PutValue('GL_QTEFACT',  QTE);
         TOBUniPiece.PutValue('GL_QTESTOCK', QTE);
         TOBUniPiece.PutValue('GL_QTERESTE', QTE);
         // --- GUINIER ---
         MTRESTE := Valeur(TOBUniPiece.GetValue('GL_MONTANTHTDEV')) + Valeur(TOBLigneCB.GetValue('GL_MONTANTHTDEV'));
         TOBUniPiece.PutValue('GL_MTRESTE',  MTRESTE);

         TOBUniPiece.Putvalue('GL_RECALCULER','X');
         end
       else
         //Insere l'article unique dans la pièce
         InsereLigneDansPiece(TOBPiece,TOBLigneCB,TOBPiece.Detail.Count);
       end
     else
       begin
       //Recherche la ligne DIM identique dans TOBPiece
       TOBDimPiece := TOBPiece.FindFirst(['GL_REFARTBARRE','GL_TYPEDIM','GL_PUHT','GL_ESCOMPTE','GL_REMISELIGNE'],
                      [TOBLigneCB.GetValue('GL_REFARTBARRE'),'DIM',TOBLigneCB.GetValue('GL_PUHT'),
                      TOBLigneCB.GetValue('GL_ESCOMPTE'),TOBLigneCB.GetValue('GL_REMISELIGNE')],False);
       if TOBDimPiece <> Nil then //Dimension trouvée
         begin
         QTE := Valeur(TOBDimPiece.GetValue('GL_QTEFACT')) + Valeur(TOBLigneCB.GetValue('GL_QTEFACT'));
         TOBDimPiece.PutValue('GL_QTEFACT', QTE);
         TOBDimPiece.PutValue('GL_QTESTOCK', QTE);
         TOBDimPiece.PutValue('GL_QTERESTE', QTE);
         // --- GUINIER ---
         MTRESTE := Valeur(TOBDimPiece.GetValue('GL_MONTANTHTDEV')) + Valeur(TOBLigneCB.GetValue('GL_MONTANTHTDEV'));
         TOBDimPiece.PutValue('GL_MTRESTE',  MTRESTE);

         TOBDimPiece.Putvalue('GL_RECALCULER','X');
         end
       else
         begin
         TOBGENPiece := TOBPiece.FindFirst(['GL_CODESDIM','GL_TYPEDIM'],
                     [TOBLigneCB.GetValue('GL_CODEARTICLE'),'GEN'],False);
         if TOBGENPiece = nil then
           begin
           //création ligne GEN puis ligne DIM
           InsereLigneGeneriqueDansPiece(TOBPiece,TOBLigneCB);
           Indice := TOBPiece.Detail.Count;
           end
         else
           begin
           DimExiste := True;
           While (TOBGENPiece <> nil) and (DimExiste) Do
             begin
             //Vérifie que la dimension n'existe pas
             DimExiste := IsDimensionExiste(TOBPiece,TOBGENPiece.GetIndex,TOBLigneCB.GetValue('GL_REFARTBARRE'));
             if DimExiste then
               TOBGENPiece := TOBPiece.FindNext(['GL_CODESDIM','GL_TYPEDIM'],
                           [TOBLigneCB.GetValue('GL_CODEARTICLE'),'GEN'],False);
             end;
           if TOBGENPiece = nil then //création ligne GEN puis ligne DIM
             begin
             InsereLigneGeneriqueDansPiece(TOBPiece,TOBLigneCB);
             Indice := TOBPiece.Detail.Count;
             end
           else Indice := IndiceGENPlusDimensions(TOBPiece,TOBGENPiece.GetIndex);
           end;
         //Insere la dimension dans la pièce au niveau de l'indice trouvé
         if Indice<>-1 then InsereLigneDansPiece(TOBPiece,TOBLigneCB,Indice);
         end;
       end;
     TOBLigneCB.Free;
     end;
end;

//Permet de regrouper les lignes de dimensions et d'ajouter une ligne générique
//si celle-ci n'existe pas
Function FusionnerLignesCodeBarre(TOBPiece : TOB): boolean;
var TOBCB : TOB;
begin
  Result:=False;
  TOBCB := ExtraireLesLignesCodeBarre(TOBPiece);
  if TOBCB.Detail.Count>0 then
    begin
    //Trie la TOB par Article pour simplifier la fusion
    TOBCB.Detail.Sort('GL_ARTICLE;GL_PUHT;GL_ESCOMPTE;GL_REMISELIGNE');
    RegroupeLignesDistinctes(TOBCB);
    InsererTOBCBDansTOBPiece(TOBPiece,TOBCB);
    Result:=True;
    end;
  TOBCB.Free;
end;

end.
