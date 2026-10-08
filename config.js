/* NovintiX Delivery Management Control Tower — deployment settings. */
window.CT_CONFIG = {
  // NovintiX Microsoft Entra ID tenant (Directory ID)
  tenant: "0ccf5041-a4b0-43b7-b7aa-5b08eda3b2b1",
  // Paste the Application (client) ID from the app registration here
  clientId: "",
  // Product Engineering SharePoint site that holds the CT_ lists
  siteUrl: "https://novintix.sharepoint.com/sites/ProductEngineering_b40832",
  // Only these email domains may sign in
  allowedDomains: ["novintix.com"],
  // Always Admin (needed for first-time setup)
  bootstrapAdmins: ["naranthera.c@novintix.com"],
  demo: false
};
