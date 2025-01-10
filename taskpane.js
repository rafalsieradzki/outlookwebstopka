Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Initialize the add-in
        document.getElementById("footer-content").innerHTML = `
            <div class="footer-familijna">
                <div>
                    <img src="https://www.familijna.pl/uploads/drive/familijna.png" alt="GRUPA FAMILIJNA" />
                </div>
                <div>
                    <span style="font-size: 14pt; color: #DF292F;">%%DisplayName%%</span><br />
                    <span style="font-size: 12pt;">%%Title%%</span><br /><br />
                    <a href="https://familijna.pl" style="color: #595959; text-decoration: none;">
                        <span style="color: #DF292F;">www.</span>familijna.pl
                    </a>
                    <span style="color: #DF292F;">email: </span>
                    <a href="mailto:%%Email%%" style="color: #595959; text-decoration: none;">%%Email%%</a><br />
                    <span style="color: #DF292F;">tel.</span> %%PhoneNumber%% <span style="color: #DF292F;">kom.</span> %%MobileNumber%%
                </div>
            </div>
        `;
    }
});