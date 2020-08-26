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
#include "Profile.h"

#include <ShObjIdl.h>

namespace TerminalApp
{
    class Jumplist
    {
    public:
        Jumplist() = default;

        void UpdateJumplist(gsl::span<const Profile> profiles);

    private:

        HRESULT _updateProfiles(winrt::com_ptr<IObjectCollection>& jumplistItems, gsl::span<const Profile> profiles);
        HRESULT _createShellLink(const std::wstring_view& name, const std::wstring_view& path, const std::wstring_view& args, winrt::com_ptr<IShellLink>& shLink);
    };
}

