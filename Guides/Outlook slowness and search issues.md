Below is a consolidated, senior-level response that ties together the key points from your conversation—covering Outlook slowness/search issues, Cached Exchange Mode considerations, OST sizing, and the new Outlook (a.k.a. "One Outlook"/"Project Monarch") limitations.

---

## 1. Outlook Slowness & Search Issues: Still Applicable in 2025?

**Yes, local indexing still matters.** Even in newer releases, Outlook’s performance in Cached Exchange Mode on Windows hinges on local indexing by Windows Search. Key factors remain:

1. **Index Rebuilds**  
   - Large or newly migrated mailboxes may experience sluggishness until indexing is complete. Rebuilding the Windows Search index or re-creating the Outlook profile can often help.
   - As a troubleshooting measure, temporarily switching Outlook from Cached to Online Mode (and testing search) will clarify whether local indexing is the bottleneck.

2. **Archive Mailbox in Cache**  
   - If the user’s Online Archive is included in Cached Exchange Mode, it can balloon the size of the local OST, potentially impacting performance (especially on older hardware or spinning disks).

3. **Item Count**  
   - Even if the OST file size is moderate, extremely high item counts (>100,000 in a single folder) degrade performance. Encourage archiving or reducing items per folder when possible.

4. **Hardware & Add-ins**  
   - Slow or outdated hardware (e.g., older HDDs vs. SSDs), certain antivirus engines scanning the OST, or heavy third-party Outlook add-ins can also contribute to sluggishness.

**Recommendation**  
- Keep Cached Exchange Mode enabled (the default) but limit the sync window (e.g., “Download 6 or 12 months of email”) to keep local OST small.  
- Ensure any needed older email remains accessible via Online Mode or OWA.  
- Encourage users to keep their primary mailbox folder sizes (and counts) in check.

---

## 2. Recommended Local PST/OST Size

- **Technical limit:** Modern Outlook versions (2010 and later) can support PST/OST files up to 50 GB.  
- **Practical recommendation:** Aim for under 20–25 GB to avoid performance hiccups, especially on typical user hardware.  
- **Very large mailboxes (>25 GB)** can still function, but short “pauses” or sync lag become more frequent as the .OST grows, particularly above 10 GB or 25 GB.

---

## 3. New Outlook (a.k.a. “One Outlook”) Missing Features & Limitations

The new Outlook for Windows—which is essentially a web-based Outlook experience packaged as a desktop app—still has several feature gaps when compared to “classic” Outlook desktop (Microsoft 365 Apps for Enterprise). Commonly cited limitations include:

1. **Online Archive Search**  
   - *Currently Missing in New Outlook:* Searching Online Archives natively is not yet fully supported. Microsoft’s roadmap indicates they plan to bring Online Archive browsing/searching to the new Outlook, but the feature has not rolled out to all users as of early 2025.  
   - *Workaround:* Use Outlook on the web or the classic Outlook desktop client to search archives until this is enabled.

2. **Rules & Automation Limitations**  
   - Certain advanced rule actions (e.g., “Reply using template,” “Play custom sound,” “Apply retention policy”) may not be available in the new Outlook.

3. **Offline Capabilities**  
   - The new Outlook is more cloud-reliant. Offline usage (especially in areas with intermittent connectivity) can be less robust than the classic client.

4. **File Format Compatibility**  
   - Saving emails in MSG or using older file formats might not be fully supported. The new Outlook often only supports EML.

5. **Customization & Add-ins**  
   - The interface is more streamlined but can feel less customizable to power users used to the classic Ribbon. Certain add-ins or COM-based customizations may not function yet in the new Outlook.

**Recommendation**  
- For organizations that rely heavily on advanced features (e.g., advanced Outlook rules, robust offline mode, or immediate access to Online Archive), the classic Outlook desktop client remains the safer bet until Microsoft closes these feature gaps in the new Outlook.  
- Pilot the new Outlook among users with simpler mail and scheduling needs. 

---

## 4. Tips to Alleviate Slowness and Improve User Experience

1. **Validate Cached Exchange Mode Defaults**  
   - By default, Outlook downloads 12 months of email. This is usually sufficient for most users without overwhelming local storage. For performance-sensitive users, reduce this to 6 or 3 months if feasible.

2. **Encourage Folder Management**  
   - High item counts per folder (>100k) can cause severe slowdowns. Train users to either create subfolders or rely more on powerful search features rather than storing everything in large single folders.

3. **Use Online Archive Properly**  
   - For large mailboxes, ensure the user’s Online Archive is enabled (in Exchange Online) so older mail can automatically move out of the primary mailbox.  
   - If the new Outlook does not yet support direct archive access, rely on Outlook Web Access (OWA) or the classic desktop Outlook client for searching archived items.

4. **Check Hardware, Updates, and Add-ins**  
   - Ensure Outlook is running on at least an SSD for large mailboxes.  
   - Keep Outlook and Windows up to date.  
   - Consider disabling any non-essential add-ins.

5. **Leverage OWA for Large Searches**  
   - Sometimes, searching large mailboxes is faster in Outlook on the web than local Outlook, as it queries the server directly rather than the local index.

---

## 5. Conclusion & Next Steps

- **Yes, local indexing nuances still apply** in modern Outlook clients, including the new Outlook for Windows—though the new client itself is heavily web-based and missing some advanced search scenarios (especially for Online Archives).  
- **Monitor mailbox/OST size and item counts** to keep the client responsive.  
- **If Online Archive access and searching is critical**, the classic Outlook desktop app or OWA remains the best workaround until Microsoft fully enables the Archive feature in the new Outlook.  
- **Keep an eye on Microsoft 365 roadmap announcements** for the timeline of new Outlook feature parity, particularly around advanced rules, offline mode improvements, and full Online Archive functionality.

By following these recommendations and staying current with Microsoft’s updates, you can mitigate most slowness or search issues and plan your eventual transition from classic Outlook to the new unified Outlook experience—once it meets all required business and user needs.