from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from typing import List, Dict, Any
from urllib.parse import quote

import pandas as pd
import requests
from dotenv import load_dotenv
from requests_ntlm import HttpNtlmAuth

# --------------------------------------------------------------------------- #
# Configuration and constants
# --------------------------------------------------------------------------- #
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
)

# Constants
LIST_TITLE = "RA4-1 Solicitud para Viajes"
DEFAULT_TIMEOUT = 30
REQUEST_HEADERS = {
    "Accept": "application/json;odata=nometadata",
    "Content-Type": "application/json",
}

# Environment variables
USERNAME = os.getenv("SP_USERNAME", "nintexinstall")
PASSWORD = os.getenv("SP_PASSWORD", "")
SITE_URL = os.getenv("SP_SITE_URL", "https://sgi.cedia.org.ec")

# --------------------------------------------------------------------------- #
# Configuration data class
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class SharePointConfig:
    """Configuration container for SharePoint connection parameters."""
    site_url: str
    username: str
    password: str
    list_title: str


# --------------------------------------------------------------------------- #
# Main connector
# --------------------------------------------------------------------------- #
class SharePointConnector:
    """Simplified connection to SharePoint Server 2019 using REST API + NTLM authentication."""

    def __init__(self, config: SharePointConfig) -> None:
        """
        Initialize the SharePoint connector.

        Parameters
        ----------
        config : SharePointConfig
            Configuration object containing connection parameters.
        """
        self._config = config
        self._session = self._create_session()

    # --------------------------- Public API -------------------------------- #
    def get_items_with_limit(self, limit: int) -> pd.DataFrame:
        """
        Retrieve up to *limit* items from the SharePoint list.

        Parameters
        ----------
        limit : int
            Maximum number of items to retrieve.

        Returns
        -------
        pd.DataFrame
            DataFrame containing the retrieved items.

        Raises
        ------
        ValueError
            If limit is less than or equal to zero.
        """
        self._validate_limit(limit)
        
        url = self._build_items_url_with_limit(limit)
        items = self._fetch_items(url)
        return self._to_dataframe(items)

    def get_all_items(self) -> pd.DataFrame:
        """
        Retrieve all items from the SharePoint list, handling internal pagination.

        Returns
        -------
        pd.DataFrame
            DataFrame containing all retrieved items.
        """
        url = self._build_items_url()
        items = self._fetch_items(url, paginate=True)
        return self._to_dataframe(items)

    # --------------------------- Private methods --------------------------- #
    @property
    def _base_api_url(self) -> str:
        """Get the base SharePoint REST API URL."""
        return f"{self._config.site_url}/_api/web"

    @property
    def _encoded_list_title(self) -> str:
        """Get the URL-encoded list title for safe use in URLs."""
        return quote(self._config.list_title)

    def _validate_limit(self, limit: int) -> None:
        """
        Validate that the limit parameter is valid.

        Parameters
        ----------
        limit : int
            The limit value to validate.

        Raises
        ------
        ValueError
            If limit is less than or equal to zero.
        """
        if limit <= 0:
            raise ValueError("Limit must be greater than zero")

    def _build_items_url(self) -> str:
        """
        Build the base URL for retrieving list items.

        Returns
        -------
        str
            The base URL for list items endpoint.
        """
        return f"{self._base_api_url}/lists/GetByTitle('{self._encoded_list_title}')/items"

    def _build_items_url_with_limit(self, limit: int) -> str:
        """
        Build the URL for retrieving list items with a limit.

        Parameters
        ----------
        limit : int
            Maximum number of items to retrieve.

        Returns
        -------
        str
            The URL with $top parameter for limiting results.
        """
        return f"{self._build_items_url()}?$top={limit}"

    def _create_session(self) -> requests.Session:
        """
        Create and configure a requests session with NTLM authentication.

        Returns
        -------
        requests.Session
            Configured session with authentication and headers.
        """
        session = requests.Session()
        session.auth = HttpNtlmAuth(self._config.username, self._config.password)
        session.headers.update(REQUEST_HEADERS)
        return session

    def _make_request(self, url: str) -> Dict[str, Any]:
        """
        Execute an HTTP request and handle errors.

        Parameters
        ----------
        url : str
            The URL to request.

        Returns
        -------
        Dict[str, Any]
            The JSON response data.

        Raises
        ------
        requests.HTTPError
            If the HTTP request fails.
        """
        logging.info("Requesting URL: %s", url)
        response = self._session.get(url, timeout=DEFAULT_TIMEOUT)
        
        try:
            response.raise_for_status()
        except requests.HTTPError as exc:
            logging.error("HTTP Error %s: %s", response.status_code, exc)
            raise

        return response.json()

    def _get_next_url(self, data: Dict[str, Any]) -> str | None:
        """
        Extract the pagination URL from the response.

        Parameters
        ----------
        data : Dict[str, Any]
            The JSON response data.

        Returns
        -------
        str | None
            The next page URL if available, None otherwise.
        """
        return (
            data.get("__next") or              # SharePoint 2010/2013/2016
            data.get("odata.nextLink") or      # nometadata
            data.get("@odata.nextLink")        # minimal/full metadata
        )

    def _fetch_items(self, url: str, paginate: bool = False) -> List[Dict[str, Any]]:
        """
        Execute the REST call and handle pagination if required.

        Parameters
        ----------
        url : str
            Initial query URL.
        paginate : bool, optional
            If True, follows pagination links until completion. Default is False.

        Returns
        -------
        List[Dict[str, Any]]
            List of retrieved items.
        """
        items: List[Dict[str, Any]] = []

        while url:
            data = self._make_request(url)
            items.extend(data.get("value", []))

            url = self._get_next_url(data) if paginate else None

        return items

    @staticmethod
    def _clean_item(item: Dict[str, Any]) -> Dict[str, Any]:
        """
        Clean an individual item by removing technical metadata.

        Parameters
        ----------
        item : Dict[str, Any]
            The item to clean.

        Returns
        -------
        Dict[str, Any]
            The cleaned item without metadata fields.
        """
        return {k: v for k, v in item.items() if not k.startswith("_")}

    @staticmethod
    def _to_dataframe(items: List[Dict[str, Any]]) -> pd.DataFrame:
        """
        Convert the list of dictionaries to DataFrame, cleaning internal metadata.

        Parameters
        ----------
        items : List[Dict[str, Any]]
            List of items to convert.

        Returns
        -------
        pd.DataFrame
            DataFrame with cleaned records.
        """
        if not items:
            logging.warning("No items were retrieved from the list.")
            return pd.DataFrame()

        cleaned_records = [SharePointConnector._clean_item(item) for item in items]
        return pd.DataFrame(cleaned_records)


# --------------------------------------------------------------------------- #
# Utility functions
# --------------------------------------------------------------------------- #
def create_config() -> SharePointConfig:
    """
    Create SharePoint configuration from environment variables.

    Returns
    -------
    SharePointConfig
        Configuration object with connection parameters.
    """
    return SharePointConfig(
        site_url=SITE_URL,
        username=USERNAME,
        password=PASSWORD,
        list_title=LIST_TITLE,
    )

def save_to_csv(dataframe: pd.DataFrame, list_title: str, filename: str = None) -> None:
    """
    Save the DataFrame to a CSV file.

    Parameters
    ----------
    dataframe : pd.DataFrame
        The DataFrame to save.
    list_title : str
        The SharePoint list title to use as filename base.
    filename : str, optional
        The output filename. If None, uses the list title as base name.
    """
    if filename is None:
        # Clean the list title to make it filesystem-safe
        safe_title = "".join(c for c in list_title if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_title = safe_title.replace(' ', '_')
        filename = f"{safe_title}.csv"
    
    dataframe.to_csv(filename, index=False)
    logging.info("Data saved to %s", filename)

def display_results(dataframe: pd.DataFrame) -> None:
    """
    Display the results in console.

    Parameters
    ----------
    dataframe : pd.DataFrame
        The DataFrame to display.
    """
    print(f"Retrieved {len(dataframe)} records.")
    print(dataframe.head())

# --------------------------------------------------------------------------- #
# Entry point (main)
# --------------------------------------------------------------------------- #
def main() -> None:
    """Main program function."""
    config = create_config()
    sp_connector = SharePointConnector(config)

    # Download without limit (handles internal pagination)
    df_all = sp_connector.get_all_items()
    
    display_results(df_all)
    save_to_csv(df_all, config.list_title)


if __name__ == "__main__":
    main()