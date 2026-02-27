import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISharePointBirthdayItem, IBirthday, LeapYearStrategy } from '../models/IBirthday';
import { processAndFilterBirthdays } from '../utils/dateUtils';

/**
 * Service for fetching birthday data from the SharePoint Verjaardagen list.
 *
 * Uses SharePoint REST API for direct list access without Graph permissions.
 * Verjaardagen list fields:
 * - Title: Naam van de persoon
 * - Geboortedatum: Geboortedatum (DateOnly)
 * - DatumEersteWerkdag: Datum eerste werkdag (DateOnly)
 */
export class BirthdayService {
  private spHttpClient: SPHttpClient;
  private siteUrl: string;
  private siteRelativeUrl: string;

  // GUIDs for the custom columns (must match elements.xml / schema.xml)
  private static readonly GEBOORTEDATUM_ID = '{3A7F9E2A-4B6C-4D8E-9F1A-2B3C4D5E6F7A}';
  private static readonly DATUM_EERSTE_WERKDAG_ID = '{7E2B5C4D-6A1F-4E9B-8C3D-1F2A3B4C5D6E}';
  private static readonly PERSOON_ID = '{2C4D8E9F-1A3B-4C5D-6E7F-8A9B0C1D2E3F}';

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
    const url = new URL(siteUrl);
    this.siteRelativeUrl = url.pathname;
  }

  /**
   * Ensures the Verjaardagen list exists. Creates it if it doesn't.
   * Silently fails if the user lacks Create Lists permission.
   */
  public async ensureList(listTitle: string): Promise<void> {
    try {
      const checkUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`;
      const checkResponse: SPHttpClientResponse = await this.spHttpClient.get(
        checkUrl,
        SPHttpClient.configurations.v1
      );
      if (checkResponse.ok) return; // List already exists
      if (checkResponse.status !== 404) return; // Unknown error, skip

      const createUrl = `${this.siteUrl}/_api/web/lists`;
      const createResponse: SPHttpClientResponse = await this.spHttpClient.post(
        createUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Content-Type': 'application/json;odata=nometadata',
            'Accept': 'application/json;odata=nometadata'
          },
          body: JSON.stringify({
            Title: listTitle,
            BaseTemplate: 100,
            Description: 'Lijst met verjaardagen en startdatums.'
          })
        }
      );

      if (createResponse.ok) {
        console.log(`BirthdayService: Lijst '${listTitle}' aangemaakt.`);
      }
    } catch (e) {
      console.warn('BirthdayService: ensureList mislukt', e);
    }
  }

  /**
   * Ensures the required columns exist in the list.
   * Creates missing columns via REST API. Silently fails if the user
   * lacks sufficient permissions (Manage Lists required).
   */
  public async ensureListColumns(listTitle: string): Promise<void> {
    const geboortedatumXml =
      `<Field ID="${BirthdayService.GEBOORTEDATUM_ID}" Type="DateTime" DisplayName="Geboortedatum" ` +
      `Required="FALSE" Format="DateOnly" StaticName="Geboortedatum" Name="Geboortedatum" />`;

    const datumEersteWerkdagXml =
      `<Field ID="${BirthdayService.DATUM_EERSTE_WERKDAG_ID}" Type="DateTime" DisplayName="Datum eerste werkdag" ` +
      `Required="FALSE" Format="DateOnly" StaticName="DatumEersteWerkdag" Name="DatumEersteWerkdag" />`;

    const persoonXml =
      `<Field ID="${BirthdayService.PERSOON_ID}" Type="User" DisplayName="Persoon" ` +
      `Required="FALSE" StaticName="Persoon" Name="Persoon" />`;

    // Each column independently — failure of one doesn't block the others
    await this.ensureColumn(listTitle, 'Geboortedatum', geboortedatumXml)
      .catch(e => console.warn('BirthdayService: Geboortedatum aanmaken mislukt', e));

    await this.ensureColumn(listTitle, 'DatumEersteWerkdag', datumEersteWerkdagXml)
      .catch(e => console.warn('BirthdayService: DatumEersteWerkdag aanmaken mislukt', e));

    await this.ensureColumn(listTitle, 'Persoon', persoonXml)
      .catch(e => console.warn('BirthdayService: Persoon aanmaken mislukt', e));
  }

  private async ensureColumn(
    listTitle: string,
    internalName: string,
    schemaXml: string
  ): Promise<void> {
    // Check if column already exists
    const checkUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/getbyinternalnameortitle('${internalName}')`;
    const checkResponse: SPHttpClientResponse = await this.spHttpClient.get(
      checkUrl,
      SPHttpClient.configurations.v1
    );
    if (checkResponse.ok) {
      return; // Already exists
    }

    const createUrl = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/fields/createfieldasxml`;
    const createResponse: SPHttpClientResponse = await this.spHttpClient.post(
      createUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Content-Type': 'application/json;odata=nometadata',
          'Accept': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({
          parameters: {
            SchemaXml: schemaXml,
            Options: 8 // AddToDefaultView
          }
        })
      }
    );

    if (!createResponse.ok) {
      const err = await createResponse.text();
      throw new Error(`Kolom '${internalName}' aanmaken mislukt: ${createResponse.status} ${err}`);
    }

    console.log(`BirthdayService: Column '${internalName}' created.`);
  }

  // Deduplicates concurrent ensureDefaultView calls across multiple webpart instances
  private static readonly _viewSetupPromises = new Map<string, Promise<void>>();

  /**
   * Sets the default view of the list to show only the three relevant columns.
   * Checks current view fields first — only modifies when the view differs from
   * the desired state. Concurrent calls for the same list are deduplicated.
   * Silently fails for users without Manage Lists permission.
   */
  public ensureDefaultView(listTitle: string): Promise<void> {
    const key = `${this.siteUrl}|${listTitle}`;
    const existing = BirthdayService._viewSetupPromises.get(key);
    if (existing) return existing;

    const cleanup = (): void => { BirthdayService._viewSetupPromises.delete(key); };
    const promise = this._doEnsureDefaultView(listTitle).then(cleanup, (e) => { cleanup(); throw e; });

    BirthdayService._viewSetupPromises.set(key, promise);
    return promise;
  }

  private async _doEnsureDefaultView(listTitle: string): Promise<void> {
    // sessionStorage gate: only run once per browser session (survives page refresh)
    let sessionKey: string | undefined;
    try {
      sessionKey = `birthdayViewOk_${this.siteUrl}_${listTitle}`;
      if (sessionStorage.getItem(sessionKey)) return;
    } catch (e) { /* sessionStorage not available, continue */ }

    try {
      const viewBase = `${this.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/defaultview`;
      const postHeaders = {
        'Content-Type': 'application/json;odata=nometadata',
        'Accept': 'application/json;odata=nometadata'
      };
      const desiredFields = ['Persoon', 'Geboortedatum', 'DatumEersteWerkdag'];

      // Read current view fields
      const fieldsResponse: SPHttpClientResponse = await this.spHttpClient.get(
        `${viewBase}/viewfields`,
        SPHttpClient.configurations.v1
      );
      if (!fieldsResponse.ok) return;

      const fieldsData = await fieldsResponse.json();

      // Items array is more reliable than SchemaXml parsing;
      // handle both minimal-metadata (array) and verbose (results wrapper) formats
      let currentFields: string[];
      if (Array.isArray(fieldsData.Items)) {
        currentFields = fieldsData.Items as string[];
      } else if (fieldsData.Items && Array.isArray(fieldsData.Items.results)) {
        currentFields = fieldsData.Items.results as string[];
      } else {
        // Fallback: parse SchemaXml
        const schemaXml: string = fieldsData.SchemaXml || '';
        currentFields = [];
        const nameRegex = /Name="([^"]+)"/g;
        let m: RegExpExecArray | null;
        // eslint-disable-next-line no-cond-assign
        while ((m = nameRegex.exec(schemaXml)) !== null) {
          currentFields.push(m[1]);
        }
      }

      // Already correct — mark session and skip modification
      const alreadyCorrect =
        currentFields.length === desiredFields.length &&
        desiredFields.every((f, i) => currentFields[i] === f);

      if (alreadyCorrect) {
        try { if (sessionKey) sessionStorage.setItem(sessionKey, '1'); } catch (e) { /* ignore */ }
        return;
      }

      // Delete all existing view fields — abort if this fails to prevent duplicates
      const deleteResponse: SPHttpClientResponse = await this.spHttpClient.post(
        `${viewBase}/viewfields/deleteallviewfields`,
        SPHttpClient.configurations.v1,
        { headers: postHeaders, body: '{}' }
      );
      if (!deleteResponse.ok) {
        console.warn('BirthdayService: deleteallviewfields mislukt', deleteResponse.status);
        return;
      }

      // Add desired fields in order
      for (const field of desiredFields) {
        await this.spHttpClient.post(
          `${viewBase}/viewfields/addviewfield('${field}')`,
          SPHttpClient.configurations.v1,
          { headers: postHeaders }
        );
      }

      try { if (sessionKey) sessionStorage.setItem(sessionKey, '1'); } catch (e) { /* ignore */ }
      console.log('BirthdayService: Default view updated.');
    } catch (e) {
      console.warn('BirthdayService: ensureDefaultView mislukt', e);
    }
  }

  /**
   * Fetch items from the SharePoint Verjaardagen list.
   *
   * @param listPath - List title or relative path (e.g., "Verjaardagen" or "Lists/Verjaardagen")
   * @returns Array of SharePoint birthday list items
   */
  public async fetchBirthdayEvents(listPath: string): Promise<ISharePointBirthdayItem[]> {
    const cleanPath = listPath.indexOf('/') === 0 ? listPath.substring(1) : listPath;
    const listRelativeUrl = cleanPath.indexOf('/') !== -1
      ? `${this.siteRelativeUrl}/${cleanPath}`
      : `${this.siteRelativeUrl}/Lists/${cleanPath}`;

    let apiUrl = `${this.siteUrl}/_api/web/GetList(@listUrl)/items`;
    apiUrl += `?@listUrl='${encodeURIComponent(listRelativeUrl)}'`;
    apiUrl += `&$select=Id,Title,Persoon/Title,Geboortedatum,DatumEersteWerkdag`;
    apiUrl += `&$expand=Persoon`;
    apiUrl += `&$orderby=Title asc`;
    apiUrl += `&$top=500`;

    console.log('BirthdayService: API URL:', apiUrl);

    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        if (response.status === 404) {
          throw new Error(
            `De lijst '${listPath}' werd niet gevonden. ` +
            `Voeg de app toe aan deze site zodat de lijst automatisch wordt aangemaakt.`
          );
        }
        if (response.status === 400) {
          throw new Error(
            `De lijst '${listPath}' mist vereiste kolommen. ` +
            `Vernieuw de pagina als site owner om de kolommen automatisch aan te maken.`
          );
        }
        const errorText = await response.text();
        throw new Error(`Fout bij ophalen van verjaardagen: ${response.status} ${response.statusText}. ${errorText}`);
      }

      const data = await response.json();
      console.log('BirthdayService: Fetched items:', data.value?.length);
      return data.value as ISharePointBirthdayItem[];
    } catch (error) {
      console.error('BirthdayService: Error fetching birthday items', error);
      throw error;
    }
  }

  /**
   * Get processed and filtered birthdays ready for display.
   * Automatically ensures required columns exist (requires Manage Lists permission on first run).
   *
   * @param listTitle - Verjaardagen list title or path
   * @param birthdayCategory - Unused (kept for API compatibility)
   * @param daysAhead - Days in future to include
   * @param maxItems - Maximum items to return
   * @param leapYearStrategy - Leap year handling
   * @param showAge - Whether to calculate age
   * @returns Processed birthday array sorted by next occurrence
   */
  public async getUpcomingBirthdays(
    listTitle: string,
    birthdayCategory: string,
    daysAhead: number,
    maxItems: number,
    leapYearStrategy: LeapYearStrategy,
    showAge: boolean
  ): Promise<IBirthday[]> {
    // Silently create list + columns + view if missing (no-op for non-owners)
    await this.ensureList(listTitle);
    await this.ensureListColumns(listTitle);
    await this.ensureDefaultView(listTitle);

    const items = await this.fetchBirthdayEvents(listTitle);
    const today = new Date();

    const result = processAndFilterBirthdays(
      items,
      today,
      daysAhead,
      maxItems,
      leapYearStrategy,
      showAge
    );

    console.log('BirthdayService: Processed birthdays:', result.length);
    return result;
  }
}
