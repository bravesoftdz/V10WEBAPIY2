{***********UNITE*************************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 02/05/2002
Modifié le ... :   /  /    
Description .. : Mise a jour des tables libre tier et du commercial d'un 
Suite ........ : document
Suite ........ : 
Suite ........ : Réservé CEGID
Mots clefs ... : LIBRE;TIERS;COMMERCIAL;DOCUMENT;CEGID
*****************************************************************}
Unit UTofCegidLibreTiersCom ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox,Vierge,M3FP,
{$IFDEF EAGLCLIENT}
      MaineAgl,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_main,
{$endIF}
      UTOF, UTOB, AGLInit;

Function  ModifCegidLibreTiersCom (Action : TActionFiche;
                                   TobP : TOB) : boolean ;


Type
  TOF_CEGIDLIBRETIERSCOM = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnClose                  ; override ;
    public
    Action : TActionFiche ;
  end ;

Implementation

var TobPieceCEGIDLibreTiersCom : TOB;

Function ModifCegidLibreTiersCom (Action : TActionFiche;
                                  TobP : TOB) : boolean ;
var Arg, Retour : string;
    NBL, NbLigne, NBPiece : integer ;
    TSql : TQuery;
begin
Result:=False;
if (TobP=Nil) then exit;
TobPieceCEGIDLibreTiersCom := TOB.Create('PIECE', Nil,-1) ;
TobPieceCEGIDLibreTiersCom.Dupliquer(TobP, True, True);

Arg:=ActionToString(taModif);

Retour:=AglLanceFiche('GC','GCCEGIDLIBRETIERS','','',Arg);
if Retour='VALIDE' then
    begin
    TobPieceCEGIDLibreTiersCom.PutValueAllFille ('GL_REPRESENTANT',
                                                 TobPieceCEGIDLibreTiersCom.GetValue ('GP_REPRESENTANT'));
    TobP.Dupliquer(TobPieceCEGIDLibreTiersCom,True,True);
    Result:=True;
    if Action = taconsult then
        begin
        TSql := OpenSQL ('SELECT COUNT(*) FROM LIGNE ' +
                            ' WHERE GL_NATUREPIECEG="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_NATUREPIECEG') +
                            '" AND GL_SOUCHE="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_SOUCHE')+
                            '" AND GL_NUMERO=' + IntToStr(TobPieceCEGIDLibreTiersCom.GetValue('GP_NUMERO')) +
                            ' AND GL_INDICEG=' + IntToStr(TobPieceCEGIDLibreTiersCom.GetValue('GP_INDICEG')), True);
        NBLigne := TSql.Fields [0].AsInteger;
        Ferme (TSql);
        beginTrans ;
        NBL := ExecuteSQL ('UPDATE LIGNE SET GL_REPRESENTANT="'+
                            TobPieceCEGIDLibreTiersCom.GetValue('GP_REPRESENTANT') +
                            '" WHERE GL_NATUREPIECEG="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_NATUREPIECEG') +
                            '" AND GL_SOUCHE="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_SOUCHE')+
                            '" AND GL_NUMERO=' + IntToStr(TobPieceCEGIDLibreTiersCom.GetValue('GP_NUMERO')) +
                            ' AND GL_INDICEG=' + IntToStr(TobPieceCEGIDLibreTiersCom.GetValue('GP_INDICEG')));
        if NBL = NBLigne then
            begin
            NBPiece := ExecuteSQL ('UPDATE PIECE SET GP_REPRESENTANT="' +
                                TobPieceCEGIDLibreTiersCom.GetValue('GP_REPRESENTANT') +
                                '", GP_LIBRETIERS1="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS1') +
                                '", GP_LIBRETIERS2="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS2') +
                                '", GP_LIBRETIERS3="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS3') +
                                '", GP_LIBRETIERS4="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS4') +
                                '", GP_LIBRETIERS5="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS5') +
                                '", GP_LIBRETIERS6="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS6') +
                                '", GP_LIBRETIERS7="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS7') +
                                '", GP_LIBRETIERS8="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS8') +
                                '", GP_LIBRETIERS9="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERS9') +
                                '", GP_LIBRETIERSA="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_LIBRETIERSA') +
                                '" WHERE GP_NATUREPIECEG="'+TobPieceCEGIDLibreTiersCom.GetValue('GP_NATUREPIECEG')+
                                '" AND GP_SOUCHE="' + TobPieceCEGIDLibreTiersCom.GetValue('GP_SOUCHE')+
                                '" AND GP_NUMERO=' + IntToStr(TobPieceCEGIDLibreTiersCom.GetValue('GP_NUMERO')) +
                                ' AND GP_INDICEG=' + IntToStr(TobPieceCEGIDLibreTiersCom.GetValue('GP_INDICEG')));
            if NBPiece <> 1 then RollBack else CommitTrans ;
            end;
        end;
    end;
TobPieceCEGIDLibreTiersCom.Free;
end;

{==============================================================================================}
{================================= Evenement de la TOF ========================================}
{==============================================================================================}
procedure TOF_CEGIDLIBRETIERSCOM.OnArgument (Arguments : String ) ;
var St : string;
    i : integer ;
begin
  Inherited ;
St:=Arguments ;
i:=Pos('ACTION=',St) ;
if i>0 then
   begin
   System.Delete(St,1,i+6) ;
   St:=uppercase(ReadTokenSt(St)) ;
   if St='CREATION' then begin Action:=taCreat ; end ;
   if St='MODIFICATION' then begin Action:=taModif ; end ;
   if St='CONSULTATION' then begin Action:=taConsult ; end ;
   end ;
TFVierge(Ecran).Retour:='' ;
end ;

procedure TOF_CEGIDLIBRETIERSCOM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CEGIDLIBRETIERSCOM.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CEGIDLIBRETIERSCOM.OnUpdate ;
begin
  Inherited ;
TobPieceCEGIDLibreTiersCom.GetEcran(Ecran);
TFVierge(Ecran).Retour:='VALIDE' ;
end ;

procedure TOF_CEGIDLIBRETIERSCOM.OnLoad ;
begin
  Inherited ;
TobPieceCEGIDLibreTiersCom.PutEcran(Ecran);
end ;

procedure TOF_CEGIDLIBRETIERSCOM.OnClose ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [TOF_CEGIDLIBRETIERSCOM] ) ;
end.
