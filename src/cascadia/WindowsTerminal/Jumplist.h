// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
//
// Module Name:
// -
//
// Abstract:
// -
//

#pragma once

#include <winrt/TerminalApp.h>
#include <ShObjIdl.h>

class Jumplist
{
public:
    Jumplist() = default;

    void UpdateJumplist();

private:

    HRESULT _updateProfiles(winrt::com_ptr<IObjectCollection>& jumplistItems, Windows::Foundation::Collections::IVectorView<Profile> profiles);
    HRESULT _createShellLink(const std::wstring_view& name, const std::wstring_view& path, const std::wstring_view& args, winrt::com_ptr<IShellLink>& shLink);
};

