/* =====================================================================
   CG HUBS - central settings, used by every app in this repository.

   THIS REPOSITORY IS PUBLIC. Put IDs and addresses here only - never
   passwords, client secrets, staff data or customer data.
   (Client IDs and the tenant ID are not secrets: every browser that
   opens a Hub already receives them. Access is controlled by Microsoft
   sign-in and SharePoint permissions, not by hiding these values.)
   ===================================================================== */

const CG_TENANT = {
  TENANT_ID: "ea9e105c-2df6-47ef-a8e3-b4aa1d788423",
  SP_HOST:   "conceptcompg.sharepoint.com"
};

/* ---------------------------------------------------------------------
   The apps. One entry per folder in this repository.

   status:
     "on"      - shown on the launcher for everyone
     "off"     - hidden from the launcher; its address shows "not available"
     "preview" - shown only to the emails in CG_PREVIEW (for testing)
   device: "mobile" (optional) - the launcher marks the tile MOBILE, for
     apps that are mainly used on a phone
   group: "<Object ID>" (optional) - only members of this Microsoft 365
     security group can see and open the app. Leave it out to let every
     Concept Group account use the app (e.g. the time-logging phone apps).
   phone: "coming-soon" (optional) - on a phone, the app shows a
     "Mobile view coming soon" screen instead of the squeezed desktop
     layout (with a "Continue anyway" link). Remove it once the app has
     its own phone layout.
   Access to a Hub is controlled by `status` and `group` here, and by
   "Assignment required" on the app in Entra (Enterprise applications).
   SharePoint permissions still decide what data a person can read or
   change.

   Every app starts as "off" and is switched to "preview" when it has
   been brought into CG_HUBS, then "on" at switch-over.
   --------------------------------------------------------------------- */
const CG_APPS = [
  { key: "training",   folder: "training/",   name: "CG Training HUB",     unit: "Concept Training Services",
    site: "/sites/training.admin", clientId: "ebbfa226-c073-4421-a02a-606f8341e4ae", status: "preview", group: "62168841-8dd9-49f6-b19d-d05f3c5933ab" /* CG HUBS – Training */, phone: "coming-soon" },

  { key: "psp",        folder: "psp/",        name: "PSP HUB",             unit: "Payroll Services PNG",
    site: "/sites/PSP_Test",       clientId: "29475230-5116-41fa-bca1-3c49c2e1dd95", status: "preview", group: "2896ef88-bd01-4f15-8b61-1f5f40af44a6" /* CG HUBS – PSP */, phone: "coming-soon" },

  { key: "csa",        folder: "csa/",        name: "CSA HUB",             unit: "Brian Hart Consulting",
    site: "/sites/Concept_CSA",    clientId: "06c6cde1-c805-4bdf-94c8-b36213b557e9", status: "preview", group: "274cdcba-c3f2-4d8d-aaee-815d42010a16" /* CG HUBS – CSA */, phone: "coming-soon" },

  { key: "sla",        folder: "sla/",        name: "SLA HUB",             unit: "Internal & Client Support",
    site: "/sites/Helpdesk",       clientId: "2489de0c-f15b-40d0-9468-a54a81bd4f34", status: "preview", group: "692052ae-b624-44c9-9a19-8b76376a0c38" /* CG HUBS – SLA */ },

  { key: "time",       folder: "time/",       name: "CG Time",             unit: "Time, attendance & payroll",
    site: "/sites/CG_Time",        clientId: "b8139965-9ca3-493e-85b4-95f799f8e8ba", status: "preview", group: "52a5230b-11d6-4e8a-b9c4-133410636660" /* CG HUBS – CG Time */, phone: "coming-soon" },

  { key: "mobile",     folder: "mobile/",     name: "Concept Time Mobile", unit: "Log consulting & support time",
    site: "/sites/CG_Time",        clientId: "ec5454e0-c794-419b-9ec9-464e9917ef3c", status: "preview", device: "mobile" },

  { key: "attendance", folder: "attendance/", name: "CG Time Attendance",  unit: "Clock in / out",
    site: "/sites/CG_Time",        clientId: "04c72756-916b-4fed-aa36-a6771ea4e8dc", status: "preview", device: "mobile" }
];

/* ---------------------------------------------------------------------
   The launcher (the CG HUBS front page). It has its own small Entra app
   registration, "CG HUBS Launcher", with only User.Read - it needs to
   know who is signed in, not read any SharePoint data.
   --------------------------------------------------------------------- */
const CG_LAUNCHER = {
  clientId: "447f8eb4-833f-4a54-868b-19d0dd821124",
  scopes: ["User.Read"]
};

// Who people should ask for access (shown in "You don't have access" messages)
const CG_ACCESS_CONTACT = "Marcus Scully (mscully@concept.com.pg)";

// Sign-in emails that can see apps set to "preview"
const CG_PREVIEW = ["mscully@concept.com.pg"];
