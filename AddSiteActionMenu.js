var customAction = {
    "Title": "My Custom Action",
    "Location": "Microsoft.SharePoint.StandardMenu",
    "Group": "SiteActions",
    "Url": "~site/_layouts/15/settings.aspx"
};

// Add the custom action
SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function() {
    var context = SP.ClientContext.get_current();
    var web = context.get_web();
    var customActions = web.get_userCustomActions();
    var newAction = customActions.add();
    newAction.set_title(customAction.Title);
    newAction.set_location(customAction.Location);
    newAction.set_group(customAction.Group);
    newAction.set_url(customAction.Url);
    newAction.update();
    context.executeQueryAsync(function() {
        console.log("Custom action added successfully!");
    },