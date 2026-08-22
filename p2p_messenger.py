# ============================================================
# P2P LAN MESSENGER
# For Attendance & Leave Tracker
#
# No SQLite
# No central server required
# Uses:
#   - UDP LAN discovery
#   - TCP peer-to-peer communication
#   - JSONL text files
#   - Local attachment storage
# ============================================================

import os
import json
import uuid
import socket
import struct
import threading
import base64
import hashlib
import time
import re

from datetime import datetime
from tkinter import (
    Toplevel,
    Frame,
    Label,
    Button,
    Entry,
    Text,
    Listbox,
    Scrollbar,
    END,
    LEFT,
    RIGHT,
    TOP,
    BOTTOM,
    BOTH,
    X,
    Y,
    W,
    SUNKEN,
    FLAT,
    SINGLE,
    filedialog,
    messagebox,
)


# ============================================================
# NETWORK SETTINGS
# ============================================================

DISCOVERY_PORT = 37020
MESSAGING_PORT = 37021

# Central persistent conversation store
SHARED_MESSAGING_ROOT = r"\\data-server\DATA\messaging_data"


DISCOVERY_MESSAGE = "ATTENDANCE_LEAVE_MESSENGER_DISCOVERY"
DISCOVERY_REQUEST = "ATTENDANCE_LEAVE_MESSENGER_DISCOVERY_REQUEST"
DISCOVERY_RESPONSE = "ATTENDANCE_LEAVE_MESSENGER_DISCOVERY_RESPONSE"

BUFFER_SIZE = 65536

SOCKET_TIMEOUT = 5

DISCOVERY_INTERVAL = 5

QUEUE_RETRY_INTERVAL = 10


# ============================================================
# STORAGE
# ============================================================

BASE_DIR = os.path.dirname(
    os.path.abspath(__file__)
)

MESSAGING_DIR = os.path.join(
    BASE_DIR,
    "messaging_data"
)


# ============================================================
# P2P MESSENGER
# ============================================================

class P2PMessenger:

    def __init__(
        self,
        root,
        username,
        users_provider,
        colors=None
    ):

        self.root = root

        self.username = str(
            username
        ).strip()

        self.users_provider = users_provider

        self.colors = colors or {}

        # ----------------------------------------------------
        # Local storage
        # ----------------------------------------------------

        self.storage_dir = os.path.join(
            MESSAGING_DIR,
            self.username
        )

        self.attachments_dir = os.path.join(
            self.storage_dir,
            "attachments"
        )

        self.messages_file = os.path.join(
            self.storage_dir,
            "messages.jsonl"
        )

        self.pending_file = os.path.join(
            self.storage_dir,
            "pending.jsonl"
        )

        # ----------------------------------------------------
        # Central persistent conversation store
        # ----------------------------------------------------
        #
        # Every workstation uses this same UNC path.  A message is
        # written here before P2P delivery is attempted, so the
        # receiver can be offline without losing the message.
        #
        self.shared_messaging_root = SHARED_MESSAGING_ROOT

        self.shared_conversations_dir = os.path.join(
            self.shared_messaging_root,
            "conversations"
        )

        self.shared_attachments_dir = os.path.join(
            self.shared_messaging_root,
            "attachments"
        )

        self.shared_pending_dir = os.path.join(
            self.shared_messaging_root,
            "pending"
        )

        try:

            os.makedirs(
                self.shared_conversations_dir,
                exist_ok=True
            )

            os.makedirs(
                self.shared_attachments_dir,
                exist_ok=True
            )

            os.makedirs(
                self.shared_pending_dir,
                exist_ok=True
            )

            print(
                "[MESSENGER] Shared messaging store:",
                self.shared_messaging_root
            )

        except Exception as e:

            print(
                "[MESSENGER] Shared store unavailable:",
                self.shared_messaging_root,
                e
            )

        os.makedirs(
            self.attachments_dir,
            exist_ok=True
        )

        # ----------------------------------------------------
        # Runtime state
        # ----------------------------------------------------

        self.running = False

        self.tcp_socket = None

        self.discovery_socket = None

        self.network_threads = []

        self.file_lock = threading.Lock()

        self.peers = {}

        self.peers_lock = threading.Lock()

        self.window = None

        self.selected_user = None

        self.chat_text = None

        self.message_entry = None

        self.contact_list = None

        self.status_label = None

        self.unread_count = {}

        # ----------------------------------------------------
        # Tkinter lifecycle state
        # ----------------------------------------------------

        self._ui_generation = 0
        self._ui_after_ids = set()
        self._ui_lock = threading.Lock()
        self._closing = False

        # ----------------------------------------------------
        # Start networking
        # ----------------------------------------------------

        self.start_network()


    # ========================================================
    # START NETWORK
    # ========================================================

    def _ui_alive(self):
        """Return True when the Tk root is still alive."""
        if self._closing:
            return False
        try:
            return bool(self.root.winfo_exists())
        except Exception:
            return False


    def _window_alive(self):
        """Return True when the Messages window still exists."""
        if not self._ui_alive():
            return False
        try:
            return (
                self.window is not None
                and self.window.winfo_exists()
            )
        except Exception:
            return False


    def _widget_alive(self, widget):
        """Safely test whether a Tkinter widget still exists."""
        if not self._ui_alive() or widget is None:
            return False
        try:
            return bool(widget.winfo_exists())
        except Exception:
            return False


    def _schedule_ui(self, callback, delay=0):
        """
        Schedule UI work on Tk's main thread.

        Network threads must never directly manipulate Tk widgets.
        Generation checking prevents callbacks from touching a window
        that has subsequently been closed/rebuilt.
        """

        if not self._ui_alive():
            return None

        generation = self._ui_generation
        after_id_holder = [None]

        def runner():
            after_id = after_id_holder[0]

            with self._ui_lock:
                if after_id in self._ui_after_ids:
                    self._ui_after_ids.remove(after_id)

            if not self._ui_alive():
                return

            if generation != self._ui_generation:
                return

            try:
                callback()
            except Exception as e:
                print(
                    "[MESSENGER UI CALLBACK ERROR]",
                    e
                )

        try:
            after_id = self.root.after(
                max(0, int(delay)),
                runner
            )

            after_id_holder[0] = after_id

            with self._ui_lock:
                self._ui_after_ids.add(after_id)

            return after_id

        except Exception as e:
            print(
                "[MESSENGER UI SCHEDULE ERROR]",
                e
            )
            return None


    def _cancel_ui_callbacks(self):
        with self._ui_lock:
            ids = list(self._ui_after_ids)
            self._ui_after_ids.clear()

        for after_id in ids:
            try:
                self.root.after_cancel(after_id)
            except Exception:
                pass


    def start_network(self):

        if self.running:
            return

        self.running = True

        # ----------------------------------------------------
        # TCP listener
        # ----------------------------------------------------

        tcp_thread = threading.Thread(
            target=self.tcp_listener,
            daemon=True
        )

        tcp_thread.start()

        self.network_threads.append(
            tcp_thread
        )

        # ----------------------------------------------------
        # UDP discovery
        # ----------------------------------------------------

        discovery_thread = threading.Thread(
            target=self.discovery_listener,
            daemon=True
        )

        discovery_thread.start()

        self.network_threads.append(
            discovery_thread
        )

        # ----------------------------------------------------
        # Broadcast presence
        # ----------------------------------------------------

        broadcast_thread = threading.Thread(
            target=self.discovery_broadcaster,
            daemon=True
        )

        broadcast_thread.start()

        self.network_threads.append(
            broadcast_thread
        )

        # ----------------------------------------------------
        # Retry pending messages
        # ----------------------------------------------------

        retry_thread = threading.Thread(
            target=self.retry_pending_loop,
            daemon=True
        )

        retry_thread.start()

        self.network_threads.append(
            retry_thread
        )


    # ========================================================
    # STOP NETWORK
    # ========================================================

    def stop_network(self):

        self.running = False
        self._closing = True
        self._ui_generation += 1
        self._cancel_ui_callbacks()

        try:
            if self.tcp_socket:
                self.tcp_socket.close()
        except Exception:
            pass

        try:
            if self.discovery_socket:
                self.discovery_socket.close()
        except Exception:
            pass

        self.tcp_socket = None
        self.discovery_socket = None


    # ========================================================
    # GET USERS
    # ========================================================

    def get_users(self):

        try:

            users = self.users_provider()

            if isinstance(users, dict):
                return users

        except Exception:
            pass

        return {}


    # ========================================================
    # GET ACTIVE USERNAMES
    # ========================================================

    def get_active_users(self):

        users = self.get_users()

        result = []

        for username, user in users.items():

            if not user.get(
                "is_active",
                True
            ):
                continue

            if username == self.username:
                continue

            result.append(
                username
            )

        return sorted(
            result
        )


    # ========================================================
    # LOCAL IP
    # ========================================================

    def get_local_ip(self):

        sock = socket.socket(
            socket.AF_INET,
            socket.SOCK_DGRAM
        )

        try:

            sock.connect(
                (
                    "8.8.8.8",
                    80
                )
            )

            ip = sock.getsockname()[0]

        except Exception:

            ip = "127.0.0.1"

        finally:

            sock.close()

        return ip


    # ========================================================
    # TCP LISTENER
    # ========================================================

    def tcp_listener(self):

        while self.running:

            sock = None

            try:

                sock = socket.socket(
                    socket.AF_INET,
                    socket.SOCK_STREAM
                )

                self.tcp_socket = sock

                sock.setsockopt(
                    socket.SOL_SOCKET,
                    socket.SO_REUSEADDR,
                    1
                )

                sock.bind(
                    (
                        "0.0.0.0",
                        MESSAGING_PORT
                    )
                )

                sock.listen(
                    50
                )

                sock.settimeout(
                    2
                )

                print(
                    f"[MESSENGER] {self.username} "
                    f"listening on TCP 0.0.0.0:{MESSAGING_PORT}"
                )

                while self.running:

                    try:
                        client_socket, address = sock.accept()
                    except socket.timeout:
                        continue
                    except OSError:
                        break

                    print(
                        "[MESSENGER] Incoming TCP connection:",
                        address
                    )

                    threading.Thread(
                        target=self.handle_peer_connection,
                        args=(client_socket, address),
                        daemon=True
                    ).start()

            except Exception as e:

                print(
                    "[MESSENGER TCP ERROR]",
                    e
                )

                if self.running:
                    time.sleep(2)

            finally:

                try:
                    if sock:
                        sock.close()
                except Exception:
                    pass

                if self.tcp_socket is sock:
                    self.tcp_socket = None

                if self.running:
                    time.sleep(1)


    # ========================================================
    # DISCOVERY LISTENER
    # ========================================================

    def discovery_listener(self):

        try:

            self.discovery_socket = socket.socket(
                socket.AF_INET,
                socket.SOCK_DGRAM
            )

            self.discovery_socket.setsockopt(
                socket.SOL_SOCKET,
                socket.SO_REUSEADDR,
                1
            )

            self.discovery_socket.setsockopt(
                socket.SOL_SOCKET,
                socket.SO_BROADCAST,
                1
            )

            self.discovery_socket.bind(
                (
                    "",
                    DISCOVERY_PORT
                )
            )

            self.discovery_socket.settimeout(
                2
            )

            print(
                f"[MESSENGER] {self.username} "
                f"listening for discovery"
            )

            while self.running:

                try:

                    data, address = (
                        self.discovery_socket.recvfrom(
                            4096
                        )
                    )

                except socket.timeout:

                    continue

                except OSError:

                    break

                try:

                    packet = json.loads(
                        data.decode(
                            "utf-8"
                        )
                    )

                except Exception:

                    continue

                if packet.get(
                    "type"
                ) != DISCOVERY_MESSAGE:

                    continue

                peer_username = packet.get(
                    "username"
                )

                peer_ip = address[0]

                peer_port = packet.get(
                    "port",
                    MESSAGING_PORT
                )

                if not peer_username:
                    continue

                if peer_username == self.username:
                    continue

                # ------------------------------------------------
                # Only recognize users that exist in application
                # ------------------------------------------------

                users = self.get_users()

                if peer_username not in users:
                    continue

                if not users[
                    peer_username
                ].get(
                    "is_active",
                    True
                ):
                    continue

                with self.peers_lock:

                    self.peers[
                        peer_username
                    ] = {
                        "ip": peer_ip,
                        "port": peer_port,
                        "last_seen": time.time()
                    }

                print(
                    "[MESSENGER] Peer discovered:",
                    peer_username,
                    f"{peer_ip}:{peer_port}"
                )

                self._schedule_ui(
                    self.refresh_contacts
                )

        except Exception as e:

            print(
                "[MESSENGER DISCOVERY ERROR]",
                e
            )


    # ========================================================
    # DISCOVERY BROADCASTER
    # ========================================================

    def discovery_broadcaster(self):

        while self.running:

            try:

                packet = {
                    "type": DISCOVERY_MESSAGE,
                    "username": self.username,
                    "port": MESSAGING_PORT
                }

                data = json.dumps(
                    packet,
                    ensure_ascii=False
                ).encode(
                    "utf-8"
                )

                addresses = [
                    "<broadcast>",
                    "255.255.255.255"
                ]

                local_ip = self.get_local_ip()

                if local_ip.count(".") == 3:

                    parts = local_ip.split(".")

                    subnet_broadcast = (
                        f"{parts[0]}.{parts[1]}.{parts[2]}.255"
                    )

                    if subnet_broadcast not in addresses:
                        addresses.append(
                            subnet_broadcast
                        )

                for broadcast_address in addresses:

                    sock = None

                    try:

                        sock = socket.socket(
                            socket.AF_INET,
                            socket.SOCK_DGRAM
                        )

                        sock.setsockopt(
                            socket.SOL_SOCKET,
                            socket.SO_BROADCAST,
                            1
                        )

                        sock.sendto(
                            data,
                            (
                                broadcast_address,
                                DISCOVERY_PORT
                            )
                        )

                    except Exception as e:

                        print(
                            "[MESSENGER BROADCAST ERROR]",
                            broadcast_address,
                            e
                        )

                    finally:

                        try:

                            if sock:
                                sock.close()

                        except Exception:
                            pass

            except Exception as e:

                print(
                    "[MESSENGER BROADCAST ERROR]",
                    e
                )

            time.sleep(
                DISCOVERY_INTERVAL
            )


    # ========================================================
    # PACKET SEND
    # ========================================================

    def send_packet(
        self,
        sock,
        header,
        payload=b""
    ):

        header = dict(
            header
        )

        header["payload_size"] = len(
            payload
        )

        header_bytes = json.dumps(
            header,
            ensure_ascii=False
        ).encode(
            "utf-8"
        )

        packet = (
            struct.pack(
                "!I",
                len(header_bytes)
            )
            + header_bytes
            + payload
        )

        sock.sendall(
            packet
        )


    # ========================================================
    # RECEIVE EXACT NUMBER OF BYTES
    # ========================================================

    def receive_exact(
        self,
        sock,
        size
    ):

        data = b""

        while len(data) < size:

            chunk = sock.recv(
                min(
                    BUFFER_SIZE,
                    size - len(data)
                )
            )

            if not chunk:
                raise ConnectionError(
                    "Connection closed."
                )

            data += chunk

        return data


    # ========================================================
    # RECEIVE PACKET
    # ========================================================

    def receive_packet(
        self,
        sock
    ):

        raw_header_length = (
            self.receive_exact(
                sock,
                4
            )
        )

        header_length = struct.unpack(
            "!I",
            raw_header_length
        )[0]

        if header_length > 10 * 1024 * 1024:
            raise ValueError(
                "Invalid header."
            )

        header_bytes = (
            self.receive_exact(
                sock,
                header_length
            )
        )

        header = json.loads(
            header_bytes.decode(
                "utf-8"
            )
        )

        payload_size = int(
            header.get(
                "payload_size",
                0
            )
        )

        if payload_size < 0:
            raise ValueError(
                "Invalid payload size."
            )

        payload = b""

        if payload_size:

            payload = (
                self.receive_exact(
                    sock,
                    payload_size
                )
            )

        return header, payload

    def get_unread_messages(self):
        """
        Return all unread messages addressed to the current user.
        """

        unread = []

        for message in self.load_messages():

            if (
                message.get("receiver")
                == self.username
                and not message.get(
                    "read",
                    False
                )
            ):

                unread.append(
                    message
                )

        return unread
    
    def get_unread_count(self):
        """Return the number of unread messages."""

        return len(
            self.get_unread_messages()
        )

    # ========================================================
    # HANDLE PEER CONNECTION
    # ========================================================

    def handle_peer_connection(
        self,
        sock,
        address
    ):

        try:

            sock.settimeout(
                SOCKET_TIMEOUT
            )

            header, payload = (
                self.receive_packet(
                    sock
                )
            )

            packet_type = header.get(
                "type"
            )

            # ------------------------------------------------
            # PING
            # ------------------------------------------------

            if packet_type == "ping":

                self.send_packet(
                    sock,
                    {
                        "type": "pong",
                        "username": self.username
                    }
                )

            # ------------------------------------------------
            # MESSAGE
            # ------------------------------------------------

            elif packet_type == "message":

                message = header.get(
                    "message"
                )

                if not isinstance(
                    message,
                    dict
                ):

                    print(
                        "[MESSENGER] Invalid message packet from",
                        address
                    )

                    return

                sender = str(
                    message.get(
                        "sender",
                        ""
                    )
                ).strip()

                receiver = str(
                    message.get(
                        "receiver",
                        ""
                    )
                ).strip()

                print(
                    "\n========================================"
                )

                print(
                    "[MESSENGER RECEIVED MESSAGE]"
                )

                print(
                    "Current user:",
                    self.username
                )

                print(
                    "Sender:",
                    sender
                )

                print(
                    "Receiver:",
                    receiver
                )

                print(
                    "Message:",
                    message.get(
                        "message",
                        ""
                    )
                )

                print(
                    "Message ID:",
                    message.get(
                        "id",
                        ""
                    )
                )

                print(
                    "Storage:",
                    self.messages_file
                )

                print(
                    "Shared storage:",
                    SHARED_MESSAGING_ROOT
                )

                print(
                    "========================================\n"
                )

                # Only accept messages addressed to this user.
                if receiver != self.username:

                    print(
                        "[MESSENGER] Ignoring message for another user:",
                        receiver
                    )

                    return

                if not sender:

                    print(
                        "[MESSENGER] Ignoring message with empty sender."
                    )

                    return

                # Store before sending ACK.  This guarantees that an
                # ACK means the message has actually been persisted.
                stored = self.store_message(
                    message
                )

                if not stored:

                    print(
                        "[MESSENGER] Failed to persist received message:",
                        message.get("id")
                    )

                    return

                # ACK only after successful persistence.
                self.send_packet(
                    sock,
                    {
                        "type": "ack",
                        "message_id": message.get(
                            "id"
                        )
                    }
                )

                self._schedule_ui(
                    lambda m=dict(message):
                        self.receive_message_ui(m)
                )

            # ------------------------------------------------
            # FILE
            # ------------------------------------------------

            elif packet_type == "file":

                message = header.get(
                    "message"
                )

                filename = (
                    message.get(
                        "attachment_name",
                        "attachment"
                    )
                    if message
                    else "attachment"
                )

                safe_name = self.safe_filename(
                    filename
                )

                unique_name = (
                    f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_"
                    f"{safe_name}"
                )

                destination = os.path.join(
                    self.attachments_dir,
                    unique_name
                )

                with open(
                    destination,
                    "wb"
                ) as file:

                    file.write(
                        payload
                    )

                message[
                    "attachment_path"
                ] = destination

                self.store_message(
                    message
                )

                self.root.after(
                    0,
                    lambda: self.receive_message_ui(
                        message
                    )
                )

                self.send_packet(
                    sock,
                    {
                        "type": "ack",
                        "message_id": message.get(
                            "id"
                        )
                    }
                )

        except Exception as e:

            print(
                "[MESSENGER CONNECTION ERROR]",
                address,
                e
            )

        finally:

            try:
                sock.close()
            except Exception:
                pass


    # ========================================================
    # SAFE FILE NAME
    # ========================================================

    def safe_filename(
        self,
        filename
    ):

        filename = os.path.basename(
            str(filename)
        )

        allowed = []

        for char in filename:

            if (
                char.isalnum()
                or char in (
                    ".",
                    "_",
                    "-",
                    " "
                )
            ):
                allowed.append(
                    char
                )

        result = "".join(
            allowed
        ).strip()

        if not result:
            result = "attachment"

        return result


    # ========================================================
    # STORE MESSAGE
    # ========================================================

    # ========================================================
    # SHARED CONVERSATION STORAGE
    # ========================================================

    def _shared_conversation_key(
        self,
        username_a,
        username_b
    ):
        """
        Generate one deterministic filename for both directions.

        Example:
            Aditya.Paikine <-> Sam
        always becomes:
            Aditya.Paikine__Sam.jsonl
        """

        names = sorted(
            [
                str(username_a).strip(),
                str(username_b).strip()
            ],
            key=str.casefold
        )

        safe_names = []

        for name in names:

            name = re.sub(
                r'[<>:"/\\\\|?*]',
                "_",
                name
            )

            safe_names.append(
                name
            )

        return "__".join(
            safe_names
        )


    def _shared_conversation_path(
        self,
        username_a,
        username_b
    ):
        return os.path.join(
            self.shared_conversations_dir,
            self._shared_conversation_key(
                username_a,
                username_b
            ) + ".jsonl"
        )


    def _ensure_shared_store(self):

        for directory in (
            self.shared_messaging_root,
            self.shared_conversations_dir,
            self.shared_attachments_dir,
            self.shared_pending_dir
        ):

            try:

                os.makedirs(
                    directory,
                    exist_ok=True
                )

            except Exception as e:

                print(
                    "[MESSENGER] Cannot access shared folder:",
                    directory,
                    e
                )

                return False

        return True


    def _append_shared_message(
        self,
        message
    ):
        """
        Append one message to the central conversation file.

        This operation happens before P2P delivery.  Therefore an
        offline recipient still gets the message when they log in.
        """

        if not self._ensure_shared_store():
            return False

        if not isinstance(
            message,
            dict
        ):
            return False

        sender = str(
            message.get(
                "sender",
                ""
            )
        ).strip()

        receiver = str(
            message.get(
                "receiver",
                ""
            )
        ).strip()

        if not sender or not receiver:
            return False

        message_id = str(
            message.get(
                "id",
                ""
            )
        ).strip()

        if not message_id:
            message_id = str(
                uuid.uuid4()
            )

            message["id"] = message_id

        path = self._shared_conversation_path(
            sender,
            receiver
        )

        # The conversation file is append-only.  Check the existing
        # IDs first so a retry after a successful delivery does not
        # duplicate the message.
        try:

            if os.path.exists(path):

                with open(
                    path,
                    "r",
                    encoding="utf-8"
                ) as file:

                    for line in file:

                        line = line.strip()

                        if not line:
                            continue

                        try:

                            existing = json.loads(
                                line
                            )

                        except Exception:

                            continue

                        if str(
                            existing.get(
                                "id",
                                ""
                            )
                        ) == message_id:

                            return True

        except Exception as e:

            print(
                "[MESSENGER] Shared conversation read error:",
                path,
                e
            )

        message.setdefault(
            "timestamp",
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        )

        message.setdefault(
            "stored_at",
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )
        )

        message.setdefault(
            "read",
            False
        )

        message.setdefault(
            "delivered",
            False
        )

        try:

            with open(
                path,
                "a",
                encoding="utf-8"
            ) as file:

                file.write(
                    json.dumps(
                        message,
                        ensure_ascii=False
                    )
                    + "\n"
                )

                file.flush()

                try:
                    os.fsync(
                        file.fileno()
                    )
                except Exception:
                    pass

            print(
                "[MESSENGER] Conversation saved:",
                path
            )

            return True

        except Exception as e:

            print(
                "[MESSENGER] Shared conversation write error:",
                path,
                e
            )

            return False


    def _load_shared_conversation(
        self,
        other_username
    ):

        if not other_username:
            return []

        path = self._shared_conversation_path(
            self.username,
            other_username
        )

        if not os.path.exists(path):
            return []

        result = []

        try:

            with open(
                path,
                "r",
                encoding="utf-8"
            ) as file:

                for line in file:

                    line = line.strip()

                    if not line:
                        continue

                    try:

                        message = json.loads(
                            line
                        )

                    except Exception:

                        continue

                    if isinstance(
                        message,
                        dict
                    ):

                        result.append(
                            message
                        )

        except Exception as e:

            print(
                "[MESSENGER] Shared conversation read error:",
                path,
                e
            )

        result.sort(
            key=lambda item: item.get(
                "timestamp",
                ""
            )
        )

        return result


    def store_message(
        self,
        message
    ):
        """
        Persist a message in the central conversation store.

        A small local copy is also retained for backward compatibility
        with older versions of the application.
        """

        if not isinstance(
            message,
            dict
        ):
            return False

        # ----------------------------------------------------
        # Central persistent copy
        # ----------------------------------------------------

        shared_ok = self._append_shared_message(
            message
        )

        # ----------------------------------------------------
        # Backward-compatible local copy
        # ----------------------------------------------------

        message_id = str(
            message.get(
                "id",
                ""
            )
        )

        existing_ids = set()

        try:

            if os.path.exists(
                self.messages_file
            ):

                with open(
                    self.messages_file,
                    "r",
                    encoding="utf-8"
                ) as file:

                    for line in file:

                        line = line.strip()

                        if not line:
                            continue

                        try:

                            item = json.loads(
                                line
                            )

                        except Exception:

                            continue

                        existing_ids.add(
                            str(
                                item.get(
                                    "id",
                                    ""
                                )
                            )
                        )

        except Exception as e:

            print(
                "[MESSENGER] Local history read error:",
                e
            )

        if message_id not in existing_ids:

            try:

                local_message = dict(
                    message
                )

                if (
                    local_message.get(
                        "receiver"
                    )
                    == self.username
                ):

                    local_message["read"] = False

                local_message.setdefault(
                    "stored_at",
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )

                with self.file_lock:

                    os.makedirs(
                        self.storage_dir,
                        exist_ok=True
                    )

                    with open(
                        self.messages_file,
                        "a",
                        encoding="utf-8"
                    ) as file:

                        file.write(
                            json.dumps(
                                local_message,
                                ensure_ascii=False
                            )
                            + "\n"
                        )

            except Exception as e:

                print(
                    "[MESSENGER] Local history write error:",
                    e
                )

        return bool(
            shared_ok
        )


    # ========================================================
    # LOAD MESSAGES
    # ========================================================

    def load_messages(self):

        messages = []

        if not os.path.exists(
            self.messages_file
        ):
            return messages

        with self.file_lock:

            try:

                with open(
                    self.messages_file,
                    "r",
                    encoding="utf-8"
                ) as file:

                    for line in file:

                        line = line.strip()

                        if not line:
                            continue

                        try:

                            messages.append(
                                json.loads(
                                    line
                                )
                            )

                        except Exception:
                            continue

            except Exception as e:

                print(
                    "[MESSENGER READ ERROR]",
                    e
                )

        return messages


    # ========================================================
    # LOAD CONVERSATION
    # ========================================================

    def load_conversation(
        self,
        other_username
    ):
        """
        Load the complete conversation from the shared UNC store.

        This is intentionally independent of peer availability.
        """

        shared_messages = self._load_shared_conversation(
            other_username
        )

        if shared_messages:
            return shared_messages

        # Fallback for conversations created by older versions.
        result = []

        for message in self.load_messages():

            sender = message.get(
                "sender"
            )

            receiver = message.get(
                "receiver"
            )

            if (
                (
                    sender == self.username
                    and receiver == other_username
                )
                or
                (
                    sender == other_username
                    and receiver == self.username
                )
            ):

                result.append(
                    message
                )

        result.sort(
            key=lambda x: x.get(
                "timestamp",
                ""
            )
        )

        return result


    # ========================================================
    # SAVE PENDING MESSAGE
    # ========================================================

    def save_pending(
        self,
        message
    ):

        with self.file_lock:

            with open(
                self.pending_file,
                "a",
                encoding="utf-8"
            ) as file:

                file.write(
                    json.dumps(
                        message,
                        ensure_ascii=False
                    )
                    + "\n"
                )


    # ========================================================
    # LOAD PENDING
    # ========================================================

    def load_pending(self):

        result = []

        if not os.path.exists(
            self.pending_file
        ):
            return result

        with self.file_lock:

            with open(
                self.pending_file,
                "r",
                encoding="utf-8"
            ) as file:

                for line in file:

                    try:

                        result.append(
                            json.loads(
                                line
                            )
                        )

                    except Exception:
                        continue

        return result


    # ========================================================
    # REWRITE PENDING
    # ========================================================

    def save_pending_list(
        self,
        messages
    ):

        with self.file_lock:

            with open(
                self.pending_file,
                "w",
                encoding="utf-8"
            ) as file:

                for message in messages:

                    file.write(
                        json.dumps(
                            message,
                            ensure_ascii=False
                        )
                        + "\n"
                    )


    # ========================================================
    # SEND MESSAGE
    # ========================================================

    def send_message(
        self,
        receiver,
        text
    ):

        receiver = str(
            receiver
        ).strip()

        text = str(
            text
        ).strip()

        if not receiver:
            return False

        if not text:
            return False

        message = {
            "id": str(
                uuid.uuid4()
            ),
            "sender": self.username,
            "receiver": receiver,
            "timestamp": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            "message": text,
            "attachment_name": "",
            "attachment_path": "",
            "read": False,
            "delivered": False
        }

        # ----------------------------------------------------
        # IMPORTANT:
        # Save to the central conversation BEFORE attempting
        # any network connection.
        # ----------------------------------------------------

        stored = self.store_message(
            message
        )

        if not stored:

            print(
                "[MESSENGER] Message was NOT saved. "
                "It will not be sent."
            )

            try:

                self._schedule_ui(
                    lambda: messagebox.showerror(
                        "Message",
                        "The message could not be saved to:\n\n"
                        f"{SHARED_MESSAGING_ROOT}\n\n"
                        "Check that the data-server share is "
                        "available and that you have write access."
                    )
                )

            except Exception:
                pass

            return False

        # ----------------------------------------------------
        # Try real-time delivery.
        # Failure here is NOT a message failure.
        # ----------------------------------------------------

        try:

            sent = self.send_to_peer(
                receiver,
                message
            )

        except Exception as e:

            print(
                "[MESSENGER] Real-time delivery failed:",
                receiver,
                e
            )

            sent = False

        if sent:

            message["delivered"] = True

            print(
                "[MESSENGER] Message delivered immediately:",
                receiver
            )

        else:

            # Keep the existing pending mechanism so that when the
            # receiver comes online, the P2P layer can still attempt
            # immediate delivery.  The conversation itself is already
            # safely stored on the server.
            try:

                self.save_pending(
                    message
                )

            except Exception as e:

                print(
                    "[MESSENGER] Pending queue error:",
                    e
                )

            print(
                "[MESSENGER] Receiver offline. "
                "Message remains in shared conversation:",
                receiver
            )

        try:

            self._schedule_ui(
                self.refresh_chat
            )

        except Exception:
            pass

        return True


    # ========================================================
    # SEND FILE
    # ========================================================

    def send_file(
        self,
        receiver,
        file_path,
        caption=""
    ):

        if not os.path.exists(
            file_path
        ):
            return False

        try:

            with open(
                file_path,
                "rb"
            ) as file:

                payload = file.read()

        except Exception as e:

            messagebox.showerror(
                "File Error",
                str(e)
            )

            return False

        filename = self.safe_filename(
            os.path.basename(
                file_path
            )
        )

        message = {
            "id": str(
                uuid.uuid4()
            ),
            "sender": self.username,
            "receiver": receiver,
            "timestamp": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),
            "message": caption,
            "attachment_name": filename,
            "attachment_path": file_path
        }

        sent = self.send_file_to_peer(
            receiver,
            message,
            payload
        )

        if sent:

            # ------------------------------------------------
            # Store sent message locally
            # ------------------------------------------------

            self.store_message(
                message
            )

        else:

            # ------------------------------------------------
            # Store metadata + base64 in pending queue
            # ------------------------------------------------

            message[
                "_pending_file_data"
            ] = base64.b64encode(
                payload
            ).decode(
                "ascii"
            )

            self.store_message(
                message
            )

            self.save_pending(
                message
            )

        self._schedule_ui(
            self.refresh_chat
        )

        return True


    # ========================================================
    # FIND PEER
    # ========================================================

    def get_peer(
        self,
        username
    ):

        with self.peers_lock:

            peer = self.peers.get(
                username
            )

            if not peer:
                return None

            if (
                time.time()
                - peer.get(
                    "last_seen",
                    0
                )
                > 20
            ):

                del self.peers[
                    username
                ]

                return None

            return dict(
                peer
            )


    # ========================================================
    # ACTIVE PEER DISCOVERY
    # ========================================================

    def discover_peer_now(
        self,
        target_username,
        wait_seconds=2.0
    ):
        """
        Immediately broadcast our presence.

        The normal discovery broadcaster runs every few seconds, but
        a user may press Send before the two applications have
        exchanged a discovery packet.  This method gives discovery
        another chance before the message is queued.
        """

        target_username = str(
            target_username
        ).strip()

        if not target_username:
            return None

        try:

            packet = {
                "type": DISCOVERY_MESSAGE,
                "username": self.username,
                "port": MESSAGING_PORT
            }

            data = json.dumps(
                packet,
                ensure_ascii=False
            ).encode(
                "utf-8"
            )

            addresses = [
                "<broadcast>",
                "255.255.255.255"
            ]

            # Also try the local subnet directed broadcast where
            # possible.  This is useful on Windows networks where
            # generic broadcast delivery can be restricted.
            local_ip = self.get_local_ip()

            if local_ip.count(".") == 3:

                parts = local_ip.split(".")

                subnet_broadcast = (
                    f"{parts[0]}.{parts[1]}.{parts[2]}.255"
                )

                if subnet_broadcast not in addresses:
                    addresses.append(
                        subnet_broadcast
                    )

            for broadcast_address in addresses:

                try:

                    sock = socket.socket(
                        socket.AF_INET,
                        socket.SOCK_DGRAM
                    )

                    sock.setsockopt(
                        socket.SOL_SOCKET,
                        socket.SO_BROADCAST,
                        1
                    )

                    sock.settimeout(
                        0.5
                    )

                    sock.sendto(
                        data,
                        (
                            broadcast_address,
                            DISCOVERY_PORT
                        )
                    )

                    sock.close()

                except Exception as e:

                    print(
                        "[MESSENGER ACTIVE DISCOVERY ERROR]",
                        broadcast_address,
                        e
                    )

            deadline = (
                time.time()
                + float(wait_seconds)
            )

            while (
                self.running
                and time.time() < deadline
            ):

                peer = self.get_peer(
                    target_username
                )

                if peer:
                    return peer

                time.sleep(
                    0.1
                )

        except Exception as e:

            print(
                "[MESSENGER DISCOVERY ERROR]",
                target_username,
                e
            )

        return self.get_peer(
            target_username
        )


    # ========================================================
    # SEND MESSAGE TO PEER
    # ========================================================

    def send_to_peer(
        self,
        receiver,
        message
    ):

        receiver = str(
            receiver
        ).strip()

        if not receiver:
            return False

        # ----------------------------------------------------
        # First use the cached discovery result.
        # ----------------------------------------------------

        peer = self.get_peer(
            receiver
        )

        # ----------------------------------------------------
        # If the peer has not been discovered yet, actively
        # broadcast discovery before putting the message into
        # the pending queue.
        # ----------------------------------------------------

        if not peer:

            print(
                "[MESSENGER] Peer not cached; "
                f"discovering {receiver}..."
            )

            peer = self.discover_peer_now(
                receiver
            )

        if not peer:

            print(
                "[MESSENGER] Peer unavailable:",
                receiver
            )

            return False

        sock = None

        try:

            print(
                "[MESSENGER] Connecting to",
                receiver,
                peer["ip"],
                peer["port"]
            )

            sock = socket.socket(
                socket.AF_INET,
                socket.SOCK_STREAM
            )

            sock.settimeout(
                SOCKET_TIMEOUT
            )

            sock.connect(
                (
                    peer["ip"],
                    int(
                        peer.get(
                            "port",
                            MESSAGING_PORT
                        )
                    )
                )
            )

            self.send_packet(
                sock,
                {
                    "type": "message",
                    "message": message
                }
            )

            response, _ = (
                self.receive_packet(
                    sock
                )
            )

            if response.get(
                "type"
            ) == "ack":

                print(
                    "[MESSENGER] Message delivered:",
                    message.get("id"),
                    "to",
                    receiver
                )

                return True

            print(
                "[MESSENGER] Receiver did not acknowledge message:",
                receiver,
                response
            )

            return False

        except Exception as e:

            print(
                "[MESSENGER SEND ERROR]",
                receiver,
                peer,
                e
            )

            return False

        finally:

            try:

                if sock:
                    sock.close()

            except Exception:
                pass


    # ========================================================
    # SEND FILE TO PEER
    # ========================================================

    def send_file_to_peer(
        self,
        receiver,
        message,
        payload
    ):

        peer = self.get_peer(
            receiver
        )

        if not peer:
            return False

        sock = None

        try:

            sock = socket.socket(
                socket.AF_INET,
                socket.SOCK_STREAM
            )

            # Larger timeout for files
            sock.settimeout(
                max(
                    SOCKET_TIMEOUT,
                    30
                )
            )

            sock.connect(
                (
                    peer["ip"],
                    peer["port"]
                )
            )

            self.send_packet(
                sock,
                {
                    "type": "file",
                    "message": message
                },
                payload
            )

            response, _ = (
                self.receive_packet(
                    sock
                )
            )

            return (
                response.get(
                    "type"
                ) == "ack"
            )

        except Exception as e:

            print(
                "[MESSENGER FILE SEND ERROR]",
                receiver,
                e
            )

            return False

        finally:

            try:

                if sock:
                    sock.close()

            except Exception:
                pass


    # ========================================================
    # RETRY PENDING MESSAGES
    # ========================================================

    def retry_pending_loop(self):

        while self.running:

            try:

                pending = self.load_pending()

                if pending:

                    remaining = []

                    for message in pending:

                        receiver = message.get(
                            "receiver"
                        )

                        # ------------------------------------------------
                        # Pending file
                        # ------------------------------------------------

                        pending_file_data = message.get(
                            "_pending_file_data"
                        )

                        if pending_file_data:

                            try:

                                payload = base64.b64decode(
                                    pending_file_data
                                )

                                sent = (
                                    self.send_file_to_peer(
                                        receiver,
                                        message,
                                        payload
                                    )
                                )

                            except Exception:

                                sent = False

                        else:

                            sent = (
                                self.send_to_peer(
                                    receiver,
                                    message
                                )
                            )

                        if not sent:

                            remaining.append(
                                message
                            )

                    self.save_pending_list(
                        remaining
                    )

            except Exception as e:

                print(
                    "[MESSENGER RETRY ERROR]",
                    e
                )

            time.sleep(
                QUEUE_RETRY_INTERVAL
            )


    # ========================================================
    # OPEN MESSENGER
    # ========================================================

    def open_window(self):
        """Open or restore the Messages window."""

        self._closing = False
        self._ui_generation += 1

        if self._window_alive():

            try:
                self.window.deiconify()
                self.window.lift()
                self.window.focus_force()
                self.rebuild_unread_counts()
                self.refresh_contacts()
            except Exception as e:
                print(
                    "[MESSENGER WINDOW ERROR]",
                    e
                )

            return

        try:

            self.window = Toplevel(
                self.root
            )

            self.window.title(
                "Messages"
            )

            self.window.geometry(
                "1000x650"
            )

            self.window.minsize(
                800,
                500
            )

            self.window.configure(
                bg=self.colors.get(
                    "background",
                    "#f5f6fa"
                )
            )

            self.window.protocol(
                "WM_DELETE_WINDOW",
                self.close_window
            )

            self.build_ui()
            self.rebuild_unread_counts()
            self.refresh_contacts()

        except Exception as e:

            print(
                "[MESSENGER OPEN WINDOW ERROR]",
                e
            )

            try:
                if self.window and self.window.winfo_exists():
                    self.window.destroy()
            except Exception:
                pass

            self.window = None
            self.contact_list = None
            self.chat_text = None
            self.message_entry = None


    # ========================================================
    # CLOSE WINDOW
    # ========================================================

    def close_window(self):
        """
        Hide the Messages window instead of destroying it.

        This is important because discovery and network threads remain
        active after the Messages window is closed.
        """

        self._ui_generation += 1
        self._cancel_ui_callbacks()

        try:

            if (
                self.window is not None
                and self.window.winfo_exists()
            ):

                self.window.withdraw()

        except Exception as e:

            print(
                "[MESSENGER CLOSE WINDOW ERROR]",
                e
            )


    # ========================================================
    # BUILD UI
    # ========================================================

    def build_ui(self):

        background = self.colors.get(
            "background",
            "#f5f6fa"
        )

        surface = self.colors.get(
            "surface",
            "#ffffff"
        )

        primary = self.colors.get(
            "primary",
            "#1f3c88"
        )

        secondary = self.colors.get(
            "secondary",
            "#3498db"
        )

        text = self.colors.get(
            "text",
            "#333333"
        )

        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        header = Frame(
            self.window,
            bg=background
        )

        header.pack(
            fill=X,
            padx=15,
            pady=10
        )

        Label(
            header,
            text="Messages",
            font=(
                "Segoe UI",
                20,
                "bold"
            ),
            bg=background,
            fg=primary
        ).pack(
            side=LEFT
        )

        self.status_label = Label(
            header,
            text="● Online",
            font=(
                "Segoe UI",
                10,
                "bold"
            ),
            bg=background,
            fg="#27ae60"
        )

        self.status_label.pack(
            side=RIGHT
        )

        # ----------------------------------------------------
        # Main area
        # ----------------------------------------------------

        main = Frame(
            self.window,
            bg=background
        )

        main.pack(
            fill=BOTH,
            expand=True,
            padx=15,
            pady=(0, 15)
        )

        # ----------------------------------------------------
        # Contacts
        # ----------------------------------------------------

        contacts_frame = Frame(
            main,
            bg=surface,
            relief=SUNKEN,
            bd=1
        )

        contacts_frame.pack(
            side=LEFT,
            fill=Y,
            padx=(0, 10)
        )

        Label(
            contacts_frame,
            text="Employees",
            font=(
                "Segoe UI",
                12,
                "bold"
            ),
            bg=surface,
            fg=text
        ).pack(
            anchor=W,
            padx=12,
            pady=10
        )

        self.contact_list = Listbox(
            contacts_frame,
            width=28,
            font=(
                "Segoe UI",
                10
            ),
            selectmode=SINGLE,
            activestyle="none"
        )

        self.contact_list.pack(
            fill=Y,
            expand=True,
            padx=8,
            pady=5
        )

        self.contact_list.bind(
            "<<ListboxSelect>>",
            self.contact_selected
        )

        # ----------------------------------------------------
        # Chat area
        # ----------------------------------------------------

        chat_frame = Frame(
            main,
            bg=surface,
            relief=SUNKEN,
            bd=1
        )

        chat_frame.pack(
            side=LEFT,
            fill=BOTH,
            expand=True
        )

        self.chat_title = Label(
            chat_frame,
            text="Select an employee",
            font=(
                "Segoe UI",
                13,
                "bold"
            ),
            bg=surface,
            fg=primary
        )

        self.chat_title.pack(
            anchor=W,
            padx=15,
            pady=10
        )

        # ----------------------------------------------------
        # Chat text
        # ----------------------------------------------------

        chat_text_frame = Frame(
            chat_frame,
            bg=surface
        )

        chat_text_frame.pack(
            fill=BOTH,
            expand=True,
            padx=10
        )

        chat_scrollbar = Scrollbar(
            chat_text_frame
        )

        chat_scrollbar.pack(
            side=RIGHT,
            fill=Y
        )

        self.chat_text = Text(
            chat_text_frame,
            font=(
                "Segoe UI",
                10
            ),
            wrap="word",
            state="disabled",
            yscrollcommand=chat_scrollbar.set
        )

        self.chat_text.pack(
            side=LEFT,
            fill=BOTH,
            expand=True
        )

        chat_scrollbar.config(
            command=self.chat_text.yview
        )

        # ----------------------------------------------------
        # Message entry
        # ----------------------------------------------------

        entry_frame = Frame(
            chat_frame,
            bg=surface
        )

        entry_frame.pack(
            fill=X,
            padx=10,
            pady=10
        )

        Button(
            entry_frame,
            text="📎",
            font=(
                "Segoe UI",
                12,
                "bold"
            ),
            bg=secondary,
            fg="white",
            relief=FLAT,
            cursor="hand2",
            padx=10,
            pady=8,
            command=self.choose_file
        ).pack(
            side=LEFT,
            padx=(0, 6)
        )

        self.message_entry = Entry(
            entry_frame,
            font=(
                "Segoe UI",
                11
            )
        )

        self.message_entry.pack(
            side=LEFT,
            fill=X,
            expand=True,
            padx=5
        )

        self.message_entry.bind(
            "<Return>",
            self.send_current_message
        )

        Button(
            entry_frame,
            text="Send",
            font=(
                "Segoe UI",
                10,
                "bold"
            ),
            bg=secondary,
            fg="white",
            relief=FLAT,
            cursor="hand2",
            padx=18,
            pady=8,
            command=self.send_current_message
        ).pack(
            side=RIGHT,
            padx=(6, 0)
        )

    def show_message_notification(
        self,
        message
    ):
        """
        Show a notification for every newly received message.
        """

        sender = str(
            message.get(
                "sender",
                "Unknown"
            )
        ).strip()

        sender_user = self.get_users().get(
            sender,
            {}
        )

        sender_name = (
            sender_user.get(
                "full_name"
            )
            or sender
        )

        text = str(
            message.get(
                "message",
                ""
            )
        ).strip()

        attachment = str(
            message.get(
                "attachment_name",
                ""
            )
        ).strip()

        if attachment:

            if text:

                preview = (
                    f"{text}\n"
                    f"📎 {attachment}"
                )

            else:

                preview = (
                    f"📎 {attachment}"
                )

        else:

            preview = text

        if not preview:

            preview = (
                "You received a new message."
            )

        if len(preview) > 200:

            preview = (
                preview[:197]
                + "..."
            )

        try:

            # -----------------------------------------------------
            # Always notify the user
            # -----------------------------------------------------

            messagebox.showinfo(
                f"💬 New message from {sender_name}",
                preview
            )

        except Exception as e:

            print(
                "[MESSENGER NOTIFICATION ERROR]",
                e
            )

    # ========================================================
    # REFRESH CONTACTS
    # ========================================================

    def refresh_contacts(self):

        if not self._widget_alive(self.contact_list):
            return

        try:

                    if not self.contact_list:
                        return

                    current = self.selected_user

                    self.contact_list.delete(
                        0,
                        END
                    )

                    active_users = self.get_active_users()

                    for username in active_users:

                        user = self.get_users().get(
                            username,
                            {}
                        )

                        full_name = (
                            user.get(
                                "full_name"
                            )
                            or username
                        )

                        online = (
                            self.get_peer(
                                username
                            )
                            is not None
                        )

                        unread = self.unread_count.get(
                            username,
                            0
                        )

                        if unread:

                            display = (
                                f"● {full_name} "
                                f"({username}) "
                                f"[{unread}]"
                            )

                        elif online:

                            display = (
                                f"● {full_name} "
                                f"({username})"
                            )

                        else:

                            display = (
                                f"○ {full_name} "
                                f"({username})"
                            )

                        self.contact_list.insert(
                            END,
                            display
                        )

                    # ----------------------------------------------------
                    # Restore selected contact
                    # ----------------------------------------------------

                    if current in active_users:

                        index = active_users.index(
                            current
                        )

                        self.contact_list.selection_set(
                            index
                        )


        except Exception as e:

            print(
                "[MESSENGER CONTACT REFRESH ERROR]",
                e
            )


    # ========================================================
    # CONTACT SELECTED
    # ========================================================

    def contact_selected(
        self,
        event=None
    ):

        if not self._widget_alive(self.contact_list):
            return

        selection = (
            self.contact_list.curselection()
        )

        if not selection:
            return

        index = selection[0]

        active_users = self.get_active_users()

        if index >= len(
            active_users
        ):
            return

        self.selected_user = (
            active_users[index]
        )

        self.mark_conversation_read(
            self.selected_user
        )

        self.unread_count[
            self.selected_user
        ] = 0

        self.chat_title.config(
            text=(
                self.get_users().get(
                    self.selected_user,
                    {}
                ).get(
                    "full_name"
                )
                or self.selected_user
            )
        )

        self.refresh_chat()

        self.refresh_contacts()


    # ========================================================
    # REFRESH CHAT
    # ========================================================

    def refresh_chat(self):

        if not self._widget_alive(self.chat_text):
            return

        if not self.selected_user:
            return

        try:

                    if not self.chat_text:
                        return

                    if not self.selected_user:
                        return

                    messages = self.load_conversation(
                        self.selected_user
                    )

                    self.chat_text.config(
                        state="normal"
                    )

                    self.chat_text.delete(
                        "1.0",
                        END
                    )

                    for message in messages:

                        sender = message.get(
                            "sender",
                            ""
                        )

                        timestamp = message.get(
                            "timestamp",
                            ""
                        )

                        text = message.get(
                            "message",
                            ""
                        )

                        attachment = message.get(
                            "attachment_name",
                            ""
                        )

                        if sender == self.username:

                            prefix = "You"

                        else:

                            prefix = (
                                self.get_users().get(
                                    sender,
                                    {}
                                ).get(
                                    "full_name"
                                )
                                or sender
                            )

                        self.chat_text.insert(
                            END,
                            f"{prefix}  {timestamp}\n",
                            "header"
                        )

                        if text:

                            self.chat_text.insert(
                                END,
                                f"{text}\n"
                            )

                        if attachment:

                            path = message.get(
                                "attachment_path",
                                ""
                            )

                            if sender == self.username:

                                self.chat_text.insert(
                                    END,
                                    f"📎 {attachment}\n"
                                )

                            elif path:

                                self.chat_text.insert(
                                    END,
                                    f"📎 {attachment} "
                                    f"[Received]\n"
                                )

                            else:

                                self.chat_text.insert(
                                    END,
                                    f"📎 {attachment}\n"
                                )

                        self.chat_text.insert(
                            END,
                            "\n"
                        )

                    self.chat_text.config(
                        state="disabled"
                    )

                    self.chat_text.see(
                        END
                    )


        except Exception as e:

            print(
                "[MESSENGER CHAT REFRESH ERROR]",
                e
            )


    # ========================================================
    # SEND CURRENT MESSAGE
    # ========================================================

    def send_current_message(self, event=None):

        if not self.selected_user:

            messagebox.showinfo(
                "Messages",
                "Please select an employee."
            )

            return "break"

        if not self._widget_alive(self.message_entry):
            return "break"

        text = (
            self.message_entry.get()
            .strip()
        )

        if not text:
            return "break"

        receiver = self.selected_user

        self.message_entry.delete(
            0,
            END
        )

        self.message_entry.focus_set()

        threading.Thread(
            target=self._send_message_background,
            args=(receiver, text),
            daemon=True
        ).start()

        return "break"


    def _send_message_background(
        self,
        receiver,
        text
    ):

        try:
            self.send_message(
                receiver,
                text
            )
        except Exception as e:
            print(
                "[MESSENGER BACKGROUND SEND ERROR]",
                receiver,
                e
            )


    # ========================================================
    # CHOOSE FILE
    # ========================================================

    def choose_file(self):

        if not self.selected_user:

            messagebox.showinfo(
                "Messages",
                "Please select an employee."
            )

            return

        file_path = filedialog.askopenfilename(
            title="Select File to Send"
        )

        if not file_path:
            return

        caption = (
            self.message_entry.get()
            .strip()
        )

        self.send_file(
            self.selected_user,
            file_path,
            caption
        )

        self.message_entry.delete(
            0,
            END
        )

    def mark_conversation_read(
        self,
        other_username
    ):

        messages = self._load_shared_conversation(
            other_username
        )

        if not messages:
            self.unread_count[
                other_username
            ] = 0
            self.refresh_contacts()
            return

        changed = False

        for message in messages:

            if (
                message.get(
                    "receiver"
                ) == self.username
                and not message.get(
                    "read",
                    False
                )
            ):

                message["read"] = True
                changed = True

        if changed:

            path = self._shared_conversation_path(
                self.username,
                other_username
            )

            temp = path + ".read.tmp"

            try:

                with open(
                    temp,
                    "w",
                    encoding="utf-8"
                ) as file:

                    for message in messages:

                        file.write(
                            json.dumps(
                                message,
                                ensure_ascii=False
                            )
                            + "\n"
                        )

                os.replace(
                    temp,
                    path
                )

            except Exception as e:

                print(
                    "[MESSENGER] Shared read-state error:",
                    e
                )

                try:

                    if os.path.exists(
                        temp
                    ):
                        os.remove(
                            temp
                        )

                except Exception:
                    pass

        self.unread_count[
            other_username
        ] = 0

        self.refresh_contacts()


    # ========================================================
    # RECEIVE MESSAGE UI
    # ========================================================

    def receive_message_ui(self, message):
        """
        Process a newly received message.

        The message is already stored by the network thread,
        but store_message() is called again safely to guarantee
        persistence.
        """

        if not isinstance(message, dict):
            return

        sender = str(
            message.get("sender", "")
        ).strip()

        receiver = str(
            message.get("receiver", "")
        ).strip()

        # ---------------------------------------------------------
        # Validate recipient
        # ---------------------------------------------------------

        if receiver != self.username:
            return

        if not sender:
            return

        # ---------------------------------------------------------
        # Permanently store the message
        # ---------------------------------------------------------

        self.store_message(message)

        # ---------------------------------------------------------
        # Rebuild unread count from disk
        # ---------------------------------------------------------

        self.rebuild_unread_counts()

        # ---------------------------------------------------------
        # Refresh currently open conversation
        # ---------------------------------------------------------

        if self.selected_user == sender:

            self.refresh_chat()

            # Because the conversation is currently open,
            # mark received messages as read.
            self.mark_conversation_read(sender)

        # ---------------------------------------------------------
        # Refresh contact list
        # ---------------------------------------------------------

        self.refresh_contacts()

        # ---------------------------------------------------------
        # Show notification
        # ---------------------------------------------------------

        self.show_message_notification(message)

        # ---------------------------------------------------------
        # Notify main attendance application
        # ---------------------------------------------------------

        try:

            self.root.event_generate(
                "<<MessengerNewMessage>>",
                when="tail"
            )

        except Exception as e:

            print(
                "[MESSENGER EVENT ERROR]",
                e
            )


    # ========================================================
    # UNREAD COUNT
    # ========================================================

    def rebuild_unread_counts(self):
        """
        Rebuild unread counts from the shared conversation files.

        This works even if the sender sent the message while this
        workstation was completely offline.
        """

        self.unread_count = {}

        try:

            if not self._ensure_shared_store():
                return self.unread_count

            users = self.get_users()

            for username in users:

                if username == self.username:
                    continue

                count = 0

                for message in self._load_shared_conversation(
                    username
                ):

                    if (
                        message.get(
                            "receiver"
                        ) == self.username
                        and not message.get(
                            "read",
                            False
                        )
                    ):

                        count += 1

                if count:

                    self.unread_count[
                        username
                    ] = count

        except Exception as e:

            print(
                "[MESSENGER] Shared unread count error:",
                e
            )

        return self.unread_count


    def update_unread_count(self):
        """
        Refresh unread counts from persistent storage.
        """

        self.rebuild_unread_counts()

        if self.contact_list:
            self.refresh_contacts()

        return self.get_unread_count()