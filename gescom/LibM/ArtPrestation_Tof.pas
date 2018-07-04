{***********UNITE*************************************************
Auteur  ...... : Didier Carret
Créé le ...... : 30/04/2002
Modifié le ... : 02/05/2002
Description .. : Source TOF des MULS : MBOPREST_MUL,
Suite ........ : MBOPRESTS5_MUL, MBOARTFI_MUL
Mots clefs ... : TOF;PREST;PRESTS5;ARTFI
*****************************************************************}
Unit ArtPrestation_Tof ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,dbtables,Mul,Fe_Main,
{$ELSE}
     eMul,Maineagl,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,
     UtilGc,UtilArticle,AglInit,M3FP,
     UTOF ;

Type
  TOF_ARTPRESTATION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Procedure LanceMulPrestation(TypeAction : string) ;
Procedure AGLLanceFichePrestation(parms:array of variant; nb: integer) ;
Procedure LanceMulArtFinancier(TypeAction : string) ;
Procedure AGLLanceFicheArtFi(parms:array of variant; nb: integer) ;

Implementation

procedure TOF_ARTPRESTATION.OnNew ;
begin
Inherited ;
end ;

procedure TOF_ARTPRESTATION.OnDelete ;
begin
Inherited ;
end ;

procedure TOF_ARTPRESTATION.OnUpdate ;
var iCol : integer ;
    stIndice,stLabel : string ;
    F : TFMul ;
    TLIB : THLabel ;
begin
Inherited ;

F:=TFMul(Ecran) ;

// Mise en place des libellés dans les colonnes
{$IFDEF EAGLCLIENT}
for iCol:=0 to F.FListe.ColCount-1 do
    BEGIN
    if copy(F.FListe.Cells[iCol,0],1,7)='Famille' then
        BEGIN  // Mise en place des libellés des familles
        stIndice:=copy(F.FListe.Cells[iCol,0],length(F.FListe.Cells[iCol,0]),1) ;
        if (stIndice>'3') and (stIndice<'9') then stLabel:='TGA2_FAMILLENIV' else stLabel:='TGA_FAMILLENIV' ;
        TLIB:=THLabel(F.FindComponent(stLabel+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.colwidths[iCol]:=0
                                 else F.Fliste.cells[iCol,0]:=TLIB.Caption ;
            END ;
        END
    else if (copy(F.FListe.Cells[iCol,0],1,19)='Statistique article') then
        BEGIN  // Mise en place des libellés des statistiques articles
        stIndice:=copy(F.FListe.Cells[iCol,0],length(F.FListe.Cells[iCol,0]),1) ;
        TLIB:=THLabel(F.FindComponent('TGA2_STATART'+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.colwidths[iCol]:=0
                                 else F.Fliste.cells[iCol,0]:=TLIB.Caption ;
            END ;
        END ;
    END ;
{$ELSE}
for iCol:=0 to F.FListe.Columns.Count-1 do
    BEGIN
    if copy(F.FListe.Columns[iCol].Title.caption,1,7)='Famille' then
        BEGIN  // Mise en place des libellés des familles
        stIndice:=copy(F.FListe.Columns[iCol].Title.caption,length(F.FListe.Columns[iCol].Title.caption),1) ;
        if (stIndice>'3') and (stIndice<'9') then stLabel:='TGA2_FAMILLENIV' else stLabel:='TGA_FAMILLENIV' ;
        TLIB:=THLabel(F.FindComponent(stLabel+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.columns[iCol].visible:=False
                                 else F.Fliste.columns[iCol].Field.DisplayLabel:=TLIB.Caption ;
            END ;
        END
    else if (copy(F.FListe.Columns[iCol].Title.caption,1,19)='Statistique article') then
        BEGIN  // Mise en place des libellés des statistiques articles
        stIndice:=copy(F.FListe.Columns[iCol].Title.caption,length(F.FListe.Columns[iCol].Title.caption),1) ;
        TLIB:=THLabel(F.FindComponent('TGA2_STATART'+stIndice)) ;
        if TLIB<>nil then
            BEGIN
            if TLIB.Caption='.-' then F.Fliste.columns[iCol].visible:=False
                                 else F.Fliste.columns[iCol].Field.DisplayLabel:=TLIB.Caption ;
            END ;
        END ;
    END ;
{$ENDIF}

end ;

procedure TOF_ARTPRESTATION.OnLoad ;
begin
Inherited ;
end ;

procedure TOF_ARTPRESTATION.OnArgument (S : String ) ;
var iChamp,i_ind : integer ;
    Critere,ChampMul,ValMul : string ;
begin
Inherited ;

if Ecran.Name='MBOARTFI_MUL' then
    BEGIN
    END else // if (Ecran.Name='MBOPREST_MUL') or (Ecran.Name='MBOPRESTS5_MUL') then
    BEGIN
    if GetPresentation=ART_ORLI then
        BEGIN
        TFMul(Ecran).Q.Manuel:=True ; // Evite l'exécution de la requête lors de la maj de Q.Liste
        TFMul(Ecran).Q.Liste:='MBOPRESTATIONS5' ;
        for iChamp:=4 to 8 do ChangeLibre2('TGA2_FAMILLENIV'+IntToStr(iChamp),Ecran) ;
        TFMul(Ecran).Q.Manuel:=False ;
        END ;

    // Mise en forme des libellés des zones familles et stat. article
    for iChamp:=1 to 3 do ChangeLibre2('TGA_FAMILLENIV'+IntToStr(iChamp),Ecran) ;

    // DCA - FQ MODE 10791 - LTAXE1 et 2 n'existent plus !
    // Mise en forme des libellés des taxes
    //for iChamp:=1 to 2 do SetControlText('LTAXE'+IntToStr(iChamp),RechDom('GCCATEGORIETAXE','TX'+IntToStr(iChamp),False)) ;
    END ;

repeat
    Critere:=uppercase(Trim(ReadTokenSt(S))) ;
    if Critere<>'' then
        BEGIN
        i_ind:=pos('=',Critere) ;
        if i_ind>0 then
            BEGIN
            ChampMul:=copy(Critere,1,i_ind-1) ;
            ValMul:=copy(Critere,i_ind+1,length(Critere)) ;
            if ChampMul='TYPEARTICLE' then SetControlText('GA_TYPEARTICLE',ValMul)
            else if ChampMul='ACTION' then SetControlText('TYPEACTION',Critere);
            END ;
        END ;
until S='';

SetControlVisible('bInsert',StringToAction(GetControlText('TYPEACTION'))<>taConsult);

end ;

procedure TOF_ARTPRESTATION.OnClose ;
begin
Inherited ;
end ;

Procedure LanceMulPrestation(TypeAction : string) ;
begin
DispatchArtMode(10,'','','TYPEARTICLE=PRE;'+TypeAction) ;  // Données de base - Articles - Prestations
end;

Procedure AGLLanceFichePrestation(parms:array of variant; nb:integer);
var Range,Lequel,Argument : string ;
begin
Range:=string(Parms[1]) ;
Lequel:=string(Parms[2]) ;
Argument:=string(Parms[3]) ;
DispatchArtMode(11,Range,Lequel,Argument) ;  // Fiche Prestations
end;

Procedure LanceMulArtFinancier(TypeAction : string) ;
begin
AGLLanceFiche('MBO','MBOARTFI_MUL','','','TYPEARTICLE=FI;'+TypeAction) ; // Paramètres - Gestion - Opérations de caisse
end;

Procedure AGLLanceFicheArtFi(parms:array of variant; nb:integer);
var Range,Lequel,Argument : string ;
begin
Range:=string(Parms[1]) ;
Lequel:=string(Parms[2]) ;
Argument:=string(Parms[3]) ;
AGLLanceFiche('MBO','MBOARTFINANCIER',Range,Lequel,Argument) ;  // Fiche Articles Financiers
end;

Initialization
RegisterClasses([TOF_ArtPrestation]) ;
RegisterAglProc('LanceFichePrestation',TRUE,3,AGLLanceFichePrestation) ;
RegisterAglProc('LanceFicheArtFi',TRUE,3,AGLLanceFicheArtFi) ;
end.

