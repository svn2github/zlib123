// Config.h
#pragma once

// TODO: Update assembly and class names here and enter 
// the correct public key token

// These strings are specific to the managed assembly that this shim will load.
static LPCWSTR szAddInAssemblyName = 
	L"OdfWordAddin, PublicKeyToken=a641c4710f1207f8";
static LPCWSTR szConnectClassName = 
	L"OdfConverter.Wordprocessing.OdfWordAddin.Connect";
static LPCWSTR szAssemblyConfigName =
	L"OdfWordAddin.dll.config";


// This is the assembly that contains the ManagedAggregator
static LPCWSTR szOdfAddinLibAssemblyName = 
	L"OdfAddinLib, PublicKeyToken=a641c4710f1207f8";