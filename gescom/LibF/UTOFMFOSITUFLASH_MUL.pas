{***********UNITE*************************************************
Auteur  ...... : O. TARCY
Créé le ...... : 13/12/2000
Modifié le ... : 23/07/2001
Description .. : Source TOF de la FICHE : SITUFLASH_MUL
Suite ........ : Choix de la journée de vente
Mots clefs ... : TOF;UTOFMFOSITUFLASH_MUL;FO
*****************************************************************}
unit UTOFMFOSITUFLASH_MUL;

interface

uses
  Classes, sysutils,
  {$IFDEF EAGLCLIENT}
  eMul,
  {$ELSE}
  Mul,
  {$ENDIF}
  UIUtil, HCtrls, HMsgBox, HDB, UTOF;

type
  TOF_MFOSITUFLASH_MUL = class(TOF)
    procedure OnArgument(S: string); override;
    procedure FOVisu;
  end;

implementation
uses
  FOConsultCaisse;

procedure TOF_MFOSITUFLASH_MUL.OnArgument(S: string);
begin
  inherited;
  TFMul(Ecran).FListe.SetFocus
end;

procedure TOF_MFOSITUFLASH_MUL.FOVisu;
var F: TFMul;
  {$IFDEF EAGLCLIENT}
  Grid: THGrid;
  {$ELSE}
  Grid: THDBGrid;
  {$ENDIF}
  nb, i, MinNumZ, MaxNumZ: Integer;
  st: string;
  erreur: Boolean;
begin
  inherited;
  F := TFMul(Ecran);
  Grid := F.FListe;
  nb := Grid.nbSelected;
  st := '';
  erreur := False;
  MinNumZ := 0;
  MaxNumZ := 0;

  if (nb = 0) and (not Grid.AllSelected) then
    PGIInfo('Il faut sélectionner au moins une clôture' + chr(13) +
      ' pour visualiser la situation flash !', 'ERREUR DANS LA SELECTION') else
    if Grid.AllSelected then
    begin
      F.Q.First;
      while not F.Q.Eof do
      begin
        if (MinNumZ = 0) or (F.Q.FindField('GJC_NUMZCAISSE').asInteger < MinNumZ) then
          MinNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
        if (MaxNumZ = 0) or (F.Q.FindField('GJC_NUMZCAISSE').asInteger > MaxNumZ) then
          MaxNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
        F.Q.Next;
      end;
      if (MinNumZ > 0) and (MaxNumZ > 0) then
        FOConsultationCaisse(GetControlText('GJC_CAISSE'), IntToStr(MinNumZ), IntToStr(MaxNumZ), FindInsidePanel, False);
      F.bSelectAll.Click;
    end else
    //vérification que les clôtures soient consécutives
    begin
      Grid.GotoLeBookmark(0);
      {$IFDEF EAGLCLIENT}
      F.Q.TQ.Seek(F.FListe.Row - 1);
      {$ENDIF}
      MinNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
      MaxNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
      for i := 1 to nb - 1 do
      begin
        Grid.GotoLeBookmark(i);
        {$IFDEF EAGLCLIENT}
        F.Q.TQ.Seek(F.FListe.Row - 1);
        {$ENDIF}
        if F.Q.FindField('GJC_NUMZCAISSE').asInteger < MinNumZ then
          MinNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger else
          if F.Q.FindField('GJC_NUMZCAISSE').asInteger > MaxNumZ then
          MaxNumZ := F.Q.FindField('GJC_NUMZCAISSE').asInteger;
      end;
      if (MinNumZ = MaxNumZ) or (MaxNumZ - MinNumZ = nb - 1) then //1 seule clôture séctionnée ou clôtures consécutives
        FOConsultationCaisse(GetControlText('GJC_CAISSE'), IntToStr(MinNumZ), IntToStr(MaxNumZ), FindInsidePanel, False) else
      begin
        erreur := True;
        PGIInfo('Les clôtures sélectionnées doivent être consécutives !', 'ERREUR DANS LA SELECTION');
      end;
      if erreur = False then Grid.ClearSelected;
    end;
end;

initialization
  registerclasses([TOF_MFOSITUFLASH_MUL]);
end.
