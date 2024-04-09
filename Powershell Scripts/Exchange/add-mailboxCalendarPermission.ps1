foreach($user in Get-Mailbox -RecipientTypeDetails UserMailbox) {

    $cal = $user.alias+":\Calendar"
    
    Add-MailboxFolderPermission -Identity $cal -User Grp.Reviewers -AccessRights Reviewer
    Add-MailboxFolderPermission -Identity $cal -User Grp.Editors -AccessRights Editor
}

