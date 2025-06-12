const core = require("@actions/core");

const fs = require("fs");
const https = require("https");
const path = require("path");
const he = require("he");
const { default: fetch } = require("node-fetch");
const { DOMParser } = require("xmldom");

const CONFIG = {
    TIMEOUT: 120000, // 2 minutes
    MS_STORE_API: "https://storeedgefd.dsx.mp.microsoft.com/v9.0",
    ENDPOINTS: {
        FE3: `https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx`,
        FE3CR: `https://fe3cr.delivery.mp.microsoft.com/ClientWebService/client.asmx`,
    },
};

const parser = new DOMParser();

async function request(url, options = {}) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), CONFIG.TIMEOUT);
    try {
        const response = await fetch(url, {
            ...options,
            signal: controller.signal,
        });
        clearTimeout(timeoutId);
        if (!response.ok) {
            throw new Error(
                `HTTP error: status: ${response.status} on "${url}"`
            );
        }
        return response;
    } catch (error) {
        clearTimeout(timeoutId);
        throw error;
    }
}

async function postXml(url, xml, unescape = false, agent = null) {
    let result = await (
        await request(url, {
            method: "POST",
            headers: {
                "Content-Type": "application/soap+xml; charset=utf-8",
            },
            body: xml,
            ...(agent ? { agent } : {})
        })
    ).text();
    if (unescape) {
        result = he.decode(result);
    }
    const xmlDoc = parser.parseFromString(result, "text/xml");

    const parseErrors = xmlDoc.getElementsByTagName("parsererror");
    if (parseErrors.length > 0) {
        throw new Error(`XML parsing error: ${parseErrors[0].textContent}`);
    }

    return xmlDoc;
}

function getNodeValue(element) {
    return element?.firstChild?.nodeValue || null;
}

function getAttributeValue(element, attrName) {
    return element?.getAttribute?.(attrName) || null;
}

async function loadCrt(url) {
    const buffer = await (await request(url)).arrayBuffer();
    return new https.Agent({
        ca: `-----BEGIN CERTIFICATE-----
${(
    Buffer.from(buffer)
        .toString("base64")
        .match(/.{1,64}/g) || []
).join("\n")}
-----END CERTIFICATE-----`,
    });
}

async function extract(productId, outputPath) {
    core.info(`🔄 Fetch necessary certificates: Microsoft Root`);
    msRootCertAgent = await loadCrt("https://www.microsoft.com/pki/certs/MicRooCerAut2011_2011_03_22.crt");
    core.info(`✅ Fetch necessary certificates: Microsoft Root`);
    core.info(`🔄 Fetch necessary certificates: Microsoft ECC Root`);
    msEccRootCertAgent = await loadCrt("https://www.microsoft.com/pkiops/certs/Microsoft%20ECC%20Product%20Root%20Certificate%20Authority%202018.crt");
    core.info(`✅ Fetch necessary certificates: Microsoft ECC Root`);
    

    if (!fs.existsSync(outputPath)) {
        await fs.promises.mkdir(outputPath, { recursive: true });
    }
    let storeResp;
    try {
        storeResp = await (
            await request(
                `${CONFIG.MS_STORE_API}/products/${productId}?market=US&locale=en-us&deviceFamily=Windows.Desktop`
            )
        ).json();
    } catch (error) {
        core.info(`❌ Product ${productId} not found`);
        throw new Error(`Product ${productId} not found`);
    }
    if (!storeResp?.Payload) {
        core.info(`❌ Product ${productId} not found`);
        throw new Error(`Product ${productId} not found`);
    }
    const sku = storeResp.Payload.Skus[0];
    if (!sku) {
        core.info(`❌ No SKU found for product ${productId}`);
        throw new Error(`No SKU found for product ${productId}`);
    }
    if (typeof sku.FulfillmentData === "string") {
        sku.FulfillmentData = JSON.parse(sku.FulfillmentData);
    }

    const fulfillment = sku.FulfillmentData;
    if (!fulfillment) {
        core.info(`❓ Cannot find fulfillment data, consider this a Win32 app`);
        throw new Error("Cannot find fulfillment data, consider this a Win32 app");
    }

    core.info(`🔄 Fetch cookie`);
    const cookieDoc = await postXml(
        CONFIG.ENDPOINTS.FE3CR,
        `<Envelope xmlns="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
<Header>
    <a:Action mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/GetCookie</a:Action>
    <a:To mustUnderstand="1">https://fe3cr.delivery.mp.microsoft.com/ClientWebService/client.asmx</a:To>
    <Security mustUnderstand="1" xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
        <WindowsUpdateTicketsToken xmlns="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization" u:id="ClientMSA">
        </WindowsUpdateTicketsToken>
    </Security>
</Header>
<Body></Body>
</Envelope>`,
        false,
        msEccRootCertAgent
    );
    core.info(`✅ Fetch cookie`);
    const cookieElements = cookieDoc.getElementsByTagName("EncryptedData");
    if (cookieElements.length === 0) {
        core.info("❌ Cannot find cookie in response");
        throw new Error("Cannot find cookie in response");
    }

    const cookie = getNodeValue(cookieElements[0]);
    if (!cookie) {
        core.info("❌ Cookie is empty");
        throw new Error("Cookie is empty");
    }

    const catId = fulfillment.WuCategoryId;
    const pkgPrefix = fulfillment.PackageFamilyName.split("_")
        .slice(0, -1)
        .join("_");

    core.info(`🔄 Fetch updates for product ${productId}`);
    let timeStamp = new Date();
    const doc = await postXml(
        CONFIG.ENDPOINTS.FE3,
        `<s:Envelope xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:s="http://www.w3.org/2003/05/soap-envelope">
<s:Header>
    <a:Action s:mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/SyncUpdates</a:Action>
    <a:MessageID>urn:uuid:175df68c-4b91-41ee-b70b-f2208c65438e</a:MessageID>
    <a:To s:mustUnderstand="1">https://fe3.delivery.mp.microsoft.com/ClientWebService/client.asmx</a:To>
    <o:Security s:mustUnderstand="1" xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
        <Timestamp xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
            <Created>${timeStamp.toISOString()}</Created>
            <Expires>${new Date(timeStamp.getTime() + 5 * 60 * 1000).toISOString()}</Expires>
        </Timestamp>
        <wuws:WindowsUpdateTicketsToken wsu:id="ClientMSA" xmlns:wsu="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd" xmlns:wuws="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization">
            <TicketType Name="MSA" Version="1.0" Policy="MBI_SSL">
                Retail
            </TicketType>
        </wuws:WindowsUpdateTicketsToken>
    </o:Security>
</s:Header>
<s:Body>
    <SyncUpdates xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
        <cookie>
            <Expiration>2045-03-11T02:02:48Z</Expiration>
            <EncryptedData>${cookie}</EncryptedData>
        </cookie>
        <parameters>
            <ExpressQuery>false</ExpressQuery>
            <InstalledNonLeafUpdateIDs>
                <int>1</int>
                <int>2</int>
                <int>3</int>
                <int>11</int>
                <int>19</int>
                <int>544</int>
                <int>549</int>
                <int>2359974</int>
                <int>2359977</int>
                <int>5169044</int>
                <int>8788830</int>
                <int>23110993</int>
                <int>23110994</int>
                <int>54341900</int>
                <int>54343656</int>
                <int>59830006</int>
                <int>59830007</int>
                <int>59830008</int>
                <int>60484010</int>
                <int>62450018</int>
                <int>62450019</int>
                <int>62450020</int>
                <int>66027979</int>
                <int>66053150</int>
                <int>97657898</int>
                <int>98822896</int>
                <int>98959022</int>
                <int>98959023</int>
                <int>98959024</int>
                <int>98959025</int>
                <int>98959026</int>
                <int>104433538</int>
                <int>104900364</int>
                <int>105489019</int>
                <int>117765322</int>
                <int>129905029</int>
                <int>130040031</int>
                <int>132387090</int>
                <int>132393049</int>
                <int>133399034</int>
                <int>138537048</int>
                <int>140377312</int>
                <int>143747671</int>
                <int>158941041</int>
                <int>158941042</int>
                <int>158941043</int>
                <int>158941044</int>
                <int>159123858</int>
                <int>159130928</int>
                <int>164836897</int>
                <int>164847386</int>
                <int>164848327</int>
                <int>164852241</int>
                <int>164852246</int>
                <int>164852252</int>
                <int>164852253</int>
            </InstalledNonLeafUpdateIDs>
            <OtherCachedUpdateIDs>
                <int>10</int>
                <int>17</int>
                <int>2359977</int>
                <int>5143990</int>
                <int>5169043</int>
                <int>5169047</int>
                <int>8806526</int>
                <int>9125350</int>
                <int>9154769</int>
                <int>10809856</int>
                <int>23110995</int>
                <int>23110996</int>
                <int>23110999</int>
                <int>23111000</int>
                <int>23111001</int>
                <int>23111002</int>
                <int>23111003</int>
                <int>23111004</int>
                <int>24513870</int>
                <int>28880263</int>
            </OtherCachedUpdateIDs>
            <SkipSoftwareSync>false</SkipSoftwareSync>
            <NeedTwoGroupOutOfScopeUpdates>true</NeedTwoGroupOutOfScopeUpdates>
            <FilterAppCategoryIds>
                <CategoryIdentifier>
                    <Id>${catId}</Id>
                </CategoryIdentifier>
            </FilterAppCategoryIds>
            <TreatAppCategoryIdsAsInstalled>true</TreatAppCategoryIdsAsInstalled>
            <AlsoPerformRegularSync>false</AlsoPerformRegularSync>
            <ComputerSpec />
            <ExtendedUpdateInfoParameters>
                <XmlUpdateFragmentTypes>
                    <XmlUpdateFragmentType>Extended</XmlUpdateFragmentType>
                </XmlUpdateFragmentTypes>
                <Locales>
                    <string>en-US</string>
                    <string>en</string>
                </Locales>
            </ExtendedUpdateInfoParameters>
            <ClientPreferredLanguages>
                <string>en-US</string>
            </ClientPreferredLanguages>
            <ProductsParameters>
                <SyncCurrentVersionOnly>false</SyncCurrentVersionOnly>
                <DeviceAttributes>
                    BranchReadinessLevel=CB;CurrentBranch=rs_prerelease;FlightRing=Retail;FlightingBranchName=external;IsFlightingEnabled=1;InstallLanguage=en-US;OSUILocale=en-US;InstallationType=Client;DeviceFamily=Windows.Desktop;
                </DeviceAttributes>
                <CallerAttributes>Interactive=1;IsSeeker=0;</CallerAttributes>
                <Products />
            </ProductsParameters>
        </parameters>
    </SyncUpdates>
</s:Body>
</s:Envelope>`,
        true,
        msRootCertAgent
    );
    core.info(`✅ Fetch updates for product ${productId}`);
    const filenames = new Map();
    const filesNodes = Array.from(doc.getElementsByTagName("Files"));

    for (const node of filesNodes) {
        try {
            const idElements =
                node.parentNode.parentNode.getElementsByTagName("ID");
            if (idElements.length === 0) continue;

            const fileId = getNodeValue(idElements[0]);
            if (!fileId || !node.firstChild) continue;

            const installerSpecificId = getAttributeValue(
                node.firstChild,
                "InstallerSpecificIdentifier"
            );
            const fileName = getAttributeValue(node.firstChild, "FileName");

            if (!installerSpecificId || !fileName) continue;

            const filename = `${installerSpecificId}_${fileName}`;
            if (filename.startsWith(pkgPrefix)) {
                filenames.set(fileId, filename);
            }
        } catch (error) {
            core.warning(`⚠️ Warning: Error processing file node: ${error.message}`);
            continue;
        }
    }

    const updates = new Map();
    const securedFragments = Array.from(
        doc.getElementsByTagName("SecuredFragment")
    );

    for (const node of securedFragments) {
        try {
            const idElements =
                node.parentNode.parentNode.parentNode.getElementsByTagName(
                    "ID"
                );
            if (idElements.length === 0) continue;

            const fileId = getNodeValue(idElements[0]);
            if (!fileId || !filenames.has(fileId)) continue;

            let updateNode = node.parentNode.parentNode.firstChild;
            while (updateNode && updateNode.nodeType !== 1) {
                updateNode = updateNode.nextSibling;
            }

            if (!updateNode) continue;

            const updateId = getAttributeValue(updateNode, "UpdateID");
            const revisionNumber = getAttributeValue(
                updateNode,
                "RevisionNumber"
            );

            if (!updateId || !revisionNumber) continue;

            updates.set(filenames.get(fileId), {
                updateId,
                revisionNumber,
            });
        } catch (error) {
            core.warning(`⚠️ Warning: Error processing secured fragment: ${error.message}`);
            continue;
        }
    }

    const results = {};
    const urlPromises = [];

    for (const [filename, { updateId, revisionNumber }] of updates) {
        urlPromises.push(fetchUrl(filename, updateId, revisionNumber));
    }

    async function fetchUrl(filename, updateId, revision) {
        core.info(`🔄 Fetch URLs for files`);
        const doc = await postXml(
            `${CONFIG.ENDPOINTS.FE3CR}/secured`,
            `<Envelope xmlns="http://www.w3.org/2003/05/soap-envelope" xmlns:a="http://www.w3.org/2005/08/addressing" xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
<Header>
<a:Action mustUnderstand="1">http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService/GetExtendedUpdateInfo2</a:Action>
<a:To mustUnderstand="1">https://fe3cr.delivery.mp.microsoft.com/ClientWebService/client.asmx/secured</a:To>
<Security mustUnderstand="1" xmlns="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
    <WindowsUpdateTicketsToken xmlns="http://schemas.microsoft.com/msus/2014/10/WindowsUpdateAuthorization" u:id="ClientMSA">
    </WindowsUpdateTicketsToken>
</Security>
</Header>
<Body>
<GetExtendedUpdateInfo2 xmlns="http://www.microsoft.com/SoftwareDistribution/Server/ClientWebService">
    <updateIDs>
        <UpdateIdentity>
            <UpdateID>${updateId}</UpdateID>
            <RevisionNumber>${revision}</RevisionNumber>
        </UpdateIdentity>
    </updateIDs>
    <infoTypes>
        <XmlUpdateFragmentType>FileUrl</XmlUpdateFragmentType>
        <XmlUpdateFragmentType>FileDecryption</XmlUpdateFragmentType>
    </infoTypes>
    <deviceAttributes>FlightRing=Retail;</deviceAttributes>
</GetExtendedUpdateInfo2>
</Body>
</Envelope>`,
            false,
            msEccRootCertAgent
        );
        core.info(`✅ Fetch URLs for files`);
        const locations = Array.from(doc.getElementsByTagName("FileLocation"));

        for (const loc of locations) {
            const urlNodes = loc.getElementsByTagName("Url");
            if (urlNodes.length === 0) continue;

            const urlValue = getNodeValue(urlNodes[0]);
            if (urlValue && urlValue.length !== 99) {

                core.info(`🔄 Pull ${filename} from ${urlValue}`);
                response = await request(urlValue);
                const fileStream = fs.createWriteStream(path.join(outputPath, filename));
                await new Promise((resolve, reject) => {
                    response.body.pipe(fileStream);
                    fileStream.on("finish", resolve);
                    fileStream.on("error", reject);
                });
                core.info(`✅ Pull ${filename} from ${urlValue}`);

                return;
            }
        }
    }

    await Promise.all(urlPromises);

    return results;
}

extract(
    core.getInput("product-id"),
    core.getInput("output-path") || process.cwd()
)
    .then(() => {
        core.setOutput("status", "success");
    })
    .catch((error) => {
        core.setFailed(`${error.name}: ${error.message}, stack:${error.stack}`);
    });
