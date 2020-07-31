#include "pch.h"
#include "Jumplist.h"

#include <shobjidl_core.h>

void Jumplist::UpdateJumplist()
{
    auto jumplist = winrt::make_self<ICustomDestinationList>();

    // Start the Jumplist edit transaction
    // removedItems represents the destinations that the user removed themselves from the jumplist.
    // It's given to us so that we respect their wishes to remove, and so we don't re-add them.
    // Currently we're only providing a static list of profiles in tasks, which won't be able
    // to be removed. (TODO: confirm this)
    uint32_t slots;
    IObjectArray* removedItems;
    if (FAILED(jumplist->BeginList(&slots, __uuidof(removedItems), reinterpret_cast<void**>(removedItems))))
    {
        return;
    }

    // Commit the edits
    if (FAILED(jumplist->CommitList()))
    {
        return;
    }
}
